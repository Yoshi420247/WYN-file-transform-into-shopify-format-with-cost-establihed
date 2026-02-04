"""
Supabase integration for the WYN-to-Shopify transformer.

Handles:
  - Recording transform runs
  - Upserting products, variants, and images into Postgres
  - Detecting diffs (new / updated / removed products) against previous runs
  - Uploading the output CSV to Supabase Storage
"""

import os
import sys
from datetime import datetime, timezone
from decimal import Decimal


def _import_supabase():
    """Lazy import of the supabase package. Returns (create_client, Client) or (None, None)."""
    try:
        from supabase import create_client, Client
        return create_client, Client
    except Exception:
        return None, None


def get_client():
    """
    Create a Supabase client from environment variables.
    Returns None if credentials are missing or supabase is unavailable.
    """
    url = os.environ.get("SUPABASE_URL", "").strip()
    key = os.environ.get("SUPABASE_SERVICE_KEY", "").strip()

    if not url or not key:
        return None

    create_client, _ = _import_supabase()
    if create_client is None:
        print("  Warning: supabase package not available, skipping sync")
        return None

    return create_client(url, key)


# ------------------------------------------------------------------
# Run tracking
# ------------------------------------------------------------------

def create_run(client: "Client", source_file: str, reference_file: str,
               triggered_by: str = "manual") -> str:
    """Insert a new transform_runs row. Returns the run UUID."""
    result = client.table("transform_runs").insert({
        "source_file": source_file,
        "reference_file": reference_file,
        "triggered_by": triggered_by,
        "status": "running",
    }).execute()
    return result.data[0]["id"]


def complete_run(client: "Client", run_id: str, stats: dict):
    """Mark a run as completed with stats."""
    client.table("transform_runs").update({
        "status": "completed",
        "total_products": stats.get("total_products", 0),
        "total_variants": stats.get("total_variants", 0),
        "total_images": stats.get("total_images", 0),
        "new_products": stats.get("new_products", 0),
        "updated_products": stats.get("updated_products", 0),
        "removed_products": stats.get("removed_products", 0),
        "output_file": stats.get("output_file", ""),
    }).eq("id", run_id).execute()


def fail_run(client: "Client", run_id: str, error: str):
    """Mark a run as failed."""
    client.table("transform_runs").update({
        "status": "failed",
        "error_message": error[:2000],
    }).eq("id", run_id).execute()


# ------------------------------------------------------------------
# Product sync with diff detection
# ------------------------------------------------------------------

def _decimal_to_float(val):
    """Convert Decimal to float for JSON serialisation."""
    if isinstance(val, Decimal):
        return float(val)
    return val


def sync_products(client: "Client", run_id: str, products: list[dict]) -> dict:
    """
    Upsert products, variants, and images into Supabase.
    Detects new, updated, and removed products vs. previous state.

    Returns stats dict with new_products, updated_products, removed_products.
    """
    # Load existing products from DB
    existing = {}
    page_size = 1000
    offset = 0
    while True:
        result = client.table("products").select(
            "id, handle, title, product_type, tags, source_url"
        ).range(offset, offset + page_size - 1).execute()
        for row in result.data:
            existing[row["handle"]] = row
        if len(result.data) < page_size:
            break
        offset += page_size

    # Load existing variants keyed by product_id
    existing_variants = {}
    offset = 0
    while True:
        result = client.table("variants").select(
            "id, product_id, sku, option_value, wyn_price, shopify_cost, variant_price"
        ).range(offset, offset + page_size - 1).execute()
        for row in result.data:
            existing_variants.setdefault(row["product_id"], []).append(row)
        if len(result.data) < page_size:
            break
        offset += page_size

    incoming_handles = set()
    new_count = 0
    updated_count = 0
    changelog_batch = []

    for product in products:
        handle = product["handle"]
        incoming_handles.add(handle)

        # Parse WYN price from first variant
        first_variant = product["variants"][0] if product["variants"] else {}
        wyn_price = _decimal_to_float(
            Decimal(first_variant["price"]) if first_variant.get("price") else None
        )
        shopify_cost = _decimal_to_float(
            Decimal(first_variant["cost"]) if first_variant.get("cost") else None
        )

        product_row = {
            "handle": handle,
            "title": product["title"],
            "body_html": product.get("body_html", ""),
            "product_type": product.get("type", ""),
            "tags": product.get("tags", ""),
            "seo_title": product.get("seo_title", ""),
            "seo_description": product.get("seo_desc", ""),
            "status": "active",
            "source_url": "",
            "wyn_category": "",
            "wyn_categories": "",
            "last_seen_run": run_id,
        }

        if handle in existing:
            # Update existing product
            ex = existing[handle]
            product_id = ex["id"]

            # Detect changes
            changes = []
            if ex.get("title") != product["title"]:
                changes.append(("title", ex.get("title", ""), product["title"]))
            if ex.get("product_type") != product.get("type", ""):
                changes.append(("product_type", ex.get("product_type", ""), product.get("type", "")))
            if ex.get("tags") != product.get("tags", ""):
                changes.append(("tags", ex.get("tags", ""), product.get("tags", "")))

            if changes:
                updated_count += 1
                for field, old, new in changes:
                    changelog_batch.append({
                        "run_id": run_id,
                        "product_id": product_id,
                        "handle": handle,
                        "change_type": "updated",
                        "field_name": field,
                        "old_value": str(old)[:1000] if old else "",
                        "new_value": str(new)[:1000] if new else "",
                    })

            # Check for price changes in variants
            old_variants = existing_variants.get(product_id, [])
            old_prices = {v["sku"]: v for v in old_variants}
            for var in product["variants"]:
                if var["sku"] in old_prices:
                    old_v = old_prices[var["sku"]]
                    old_price = old_v.get("variant_price")
                    new_price = float(var["price"]) if var.get("price") else None
                    if old_price is not None and new_price is not None and abs(float(old_price) - new_price) > 0.005:
                        changelog_batch.append({
                            "run_id": run_id,
                            "product_id": product_id,
                            "handle": handle,
                            "change_type": "price_change",
                            "field_name": f"variant_price ({var['sku']})",
                            "old_value": str(old_price),
                            "new_value": str(new_price),
                        })

            product_row["updated_at"] = datetime.now(timezone.utc).isoformat()
            client.table("products").update(product_row).eq("id", product_id).execute()

        else:
            # New product
            new_count += 1
            product_row["first_seen_run"] = run_id
            result = client.table("products").insert(product_row).execute()
            product_id = result.data[0]["id"]

            changelog_batch.append({
                "run_id": run_id,
                "product_id": product_id,
                "handle": handle,
                "change_type": "new",
            })

        # Upsert variants
        # Delete old variants and re-insert (simpler than individual upsert)
        client.table("variants").delete().eq("product_id", product_id).execute()
        for var in product["variants"]:
            var_price = _decimal_to_float(Decimal(var["price"])) if var.get("price") else None
            var_cost = _decimal_to_float(Decimal(var["cost"])) if var.get("cost") else None
            client.table("variants").insert({
                "product_id": product_id,
                "sku": var["sku"],
                "option_name": var.get("option_name", "Title"),
                "option_value": var.get("option_value", "Default Title"),
                "wyn_price": var_price,
                "shopify_cost": var_cost,
                "variant_price": var_price,
            }).execute()

        # Upsert images
        client.table("product_images").delete().eq("product_id", product_id).execute()
        for pos, img_url in enumerate(product.get("images", []), start=1):
            client.table("product_images").insert({
                "product_id": product_id,
                "image_url": img_url,
                "position": pos,
                "alt_text": product["title"],
            }).execute()

    # Detect removed products (in DB but not in this catalogue)
    removed_count = 0
    for handle, ex in existing.items():
        if handle not in incoming_handles:
            removed_count += 1
            changelog_batch.append({
                "run_id": run_id,
                "product_id": ex["id"],
                "handle": handle,
                "change_type": "removed",
            })

    # Batch insert changelog
    if changelog_batch:
        # Insert in chunks to avoid payload limits
        chunk_size = 100
        for i in range(0, len(changelog_batch), chunk_size):
            chunk = changelog_batch[i:i + chunk_size]
            client.table("product_changelog").insert(chunk).execute()

    return {
        "new_products": new_count,
        "updated_products": updated_count,
        "removed_products": removed_count,
    }


# ------------------------------------------------------------------
# Storage upload
# ------------------------------------------------------------------

def upload_csv(client: "Client", run_id: str, csv_path: str) -> str:
    """
    Upload the output CSV to Supabase Storage.
    Returns the storage path.
    """
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    storage_path = f"runs/{run_id}/{timestamp}_shopify_import.csv"

    with open(csv_path, "rb") as f:
        client.storage.from_("transform-outputs").upload(
            path=storage_path,
            file=f,
            file_options={"content-type": "text/csv"},
        )

    return storage_path


# ------------------------------------------------------------------
# Public API
# ------------------------------------------------------------------

def run_sync(products: list[dict], csv_path: str,
             source_file: str, reference_file: str,
             triggered_by: str = "manual") -> bool:
    """
    Full sync pipeline. Returns True if sync succeeded, False if skipped/failed.
    """
    client = get_client()
    if client is None:
        print("\n  Supabase sync: skipped (no SUPABASE_URL / SUPABASE_SERVICE_KEY)")
        return False

    run_id = None
    try:
        print("\n  Supabase sync: starting...")

        # 1. Create run record
        run_id = create_run(client, source_file, reference_file, triggered_by)
        print(f"    Run ID: {run_id}")

        # 2. Sync products, variants, images + detect diffs
        print("    Syncing products to database...")
        diff_stats = sync_products(client, run_id, products)
        print(f"    New: {diff_stats['new_products']}, "
              f"Updated: {diff_stats['updated_products']}, "
              f"Removed: {diff_stats['removed_products']}")

        # 3. Upload CSV to storage
        print("    Uploading CSV to storage...")
        storage_path = upload_csv(client, run_id, csv_path)
        print(f"    Stored at: {storage_path}")

        # 4. Complete the run
        total_variants = sum(len(p["variants"]) for p in products)
        total_images = sum(len(p["images"]) for p in products)
        complete_run(client, run_id, {
            "total_products": len(products),
            "total_variants": total_variants,
            "total_images": total_images,
            "output_file": storage_path,
            **diff_stats,
        })
        print("    Supabase sync: complete")
        return True

    except Exception as e:
        print(f"    Supabase sync error: {e}")
        if run_id:
            try:
                fail_run(client, run_id, str(e))
            except Exception:
                pass
        return False
