#!/usr/bin/env python3
"""
WYN Catalogue → Shopify Product Import CSV Transformer

Reads a WYN wholesale catalogue (.xlsx) and a Shopify reference export (.csv),
then produces a Shopify-ready import CSV that:
  - Matches the exact schema, column order, and formatting of the reference CSV
  - Preserves all WYN fields (unmapped fields stored as HTML metadata)
  - Converts WYN prices to Shopify "Cost per item" using tiered multipliers
  - Handles variants, images, deduplication, and SKU generation
"""

import argparse
import csv
import html
import math
import re
import sys
from collections import OrderedDict
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from urllib.parse import unquote

import openpyxl

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

SHOPIFY_HEADERS = [
    "Handle", "Title", "Body (HTML)", "Vendor", "Product Category", "Type",
    "Tags", "Published", "Option1 Name", "Option1 Value", "Option1 Linked To",
    "Option2 Name", "Option2 Value", "Option2 Linked To", "Option3 Name",
    "Option3 Value", "Option3 Linked To", "Variant SKU", "Variant Grams",
    "Variant Inventory Tracker", "Variant Inventory Policy",
    "Variant Fulfillment Service", "Variant Price", "Variant Compare At Price",
    "Variant Requires Shipping", "Variant Taxable",
    "Unit Price Total Measure", "Unit Price Total Measure Unit",
    "Unit Price Base Measure", "Unit Price Base Measure Unit",
    "Variant Barcode", "Image Src", "Image Position", "Image Alt Text",
    "Gift Card", "SEO Title", "SEO Description",
    "Google Shopping / Google Product Category", "Google Shopping / Gender",
    "Google Shopping / Age Group", "Google Shopping / MPN",
    "Google Shopping / Condition", "Google Shopping / Custom Product",
    "Google Shopping / Custom Label 0", "Google Shopping / Custom Label 1",
    "Google Shopping / Custom Label 2", "Google Shopping / Custom Label 3",
    "Google Shopping / Custom Label 4",
    "Google: Custom Product (product.metafields.mm-google-shopping.custom_product)",
    "Color (product.metafields.shopify.color-pattern)",
    "Material (product.metafields.shopify.material)",
    "Variant Image", "Variant Weight Unit", "Variant Tax Code",
    "Cost per item", "Status",
]

WYN_HEADERS = [
    "Categories_URL", "URL", "Product_URL", "Name", "Category", "Category_URL",
    "Categories", "Price", "Description", "Description_HTML", "SKU",
    "Main_Image", "Stock", "Variants", "Variant_01", "Variant_02",
    "Variant_03", "Variant_04", "Variant_05", "Variant_06", "Variant_07",
    "Variant_08", "All_Images",
]

PRICING_TIERS = [
    (Decimal("0.50"),  Decimal("4.00"),   Decimal("2.5")),
    (Decimal("4.01"),  Decimal("20.00"),  Decimal("2.0")),
    (Decimal("20.01"), Decimal("40.00"),  Decimal("1.8")),
    (Decimal("40.01"), Decimal("100.00"), Decimal("1.6")),
    (Decimal("100.01"), Decimal("200.00"), Decimal("1.5")),
    (Decimal("200.01"), Decimal("999999999"), Decimal("1.4")),
]

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def clean(val):
    """Normalise a cell value to a stripped string or empty string."""
    if val is None:
        return ""
    s = str(val).strip()
    if s in ("None", "NaN", "null", "N/A"):
        return ""
    return s


def parse_price(raw: str) -> Decimal | None:
    """Convert '$10.00' or '$6,000.00' to Decimal, or None."""
    raw = clean(raw)
    if not raw:
        return None
    raw = raw.replace("$", "").replace(",", "").strip()
    try:
        return Decimal(raw)
    except Exception:
        return None


def format_price(d: Decimal) -> str:
    """Format Decimal to 2-decimal string."""
    return str(d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def compute_cost(price: Decimal) -> str:
    """Apply tiered multiplier to WYN price → Shopify Cost per item."""
    for lo, hi, mult in PRICING_TIERS:
        if lo <= price <= hi:
            result = price * mult
            return format_price(result)
    # Fallback for prices below 0.50 — use 2.5x
    if price < Decimal("0.50") and price > 0:
        return format_price(price * Decimal("2.5"))
    return ""


def slugify(text: str) -> str:
    """Create a URL-safe handle from text."""
    s = text.lower()
    s = re.sub(r"[^a-z0-9\-]", "-", s)
    s = re.sub(r"-{2,}", "-", s)
    s = s.strip("-")
    return s


def url_to_slug(url: str) -> str:
    """Extract slug from WYN product URL."""
    url = url.strip().rstrip("/")
    slug = url.rsplit("/", 1)[-1]
    slug = unquote(slug)
    return slugify(slug)


def normalise_url(url: str) -> str:
    """Normalise URL for deduplication."""
    return clean(url).rstrip("/").lower()


def parse_images(main_image: str, all_images: str) -> list[str]:
    """Parse and deduplicate image URLs, main image first."""
    images = []
    seen = set()

    def add(url):
        url = url.replace("_x000d_", "").replace("\r", "").strip()
        if url and url not in seen:
            seen.add(url)
            images.append(url)

    if main_image:
        add(clean(main_image))

    if all_images:
        raw = clean(all_images)
        for part in raw.split(","):
            part = part.strip()
            if part:
                add(part)

    return images


def parse_variants(variants_raw: str, variant_cols: list[str]) -> tuple[str, list[str]]:
    """
    Parse the Variants field and Variant_01..08 columns.
    Returns (option_name, [variant_values]).
    """
    option_name = ""
    values = []

    if variants_raw:
        # The Variants field starts with the option name (Size, QTY, Pack QTY)
        # followed by whitespace/tabs/newlines and then comma-separated values
        text = variants_raw.strip()

        # Extract option name — it's the first meaningful token before tab/newline
        # Known option names: Size, QTY, Pack QTY
        for known in ["Pack QTY", "Size", "QTY"]:
            if text.startswith(known):
                option_name = known
                text = text[len(known):]
                break

        if not option_name:
            # Fallback: first line/word
            first_line = text.split("\n")[0].split("\t")[0].strip()
            if first_line and not any(c.isdigit() for c in first_line):
                option_name = first_line
                text = text[len(first_line):]

        # Parse remaining text for values
        # Replace tabs and newlines with spaces, then split on commas
        text = text.replace("\t", " ").replace("\n", " ").replace("\r", " ")
        parts = text.split(",")
        for p in parts:
            p = p.strip()
            if p:
                values.append(p)

    # Union with Variant_01..08
    existing = set(values)
    for vc in variant_cols:
        vc = clean(vc)
        if vc and vc not in existing:
            values.append(vc)
            existing.add(vc)

    return option_name, values


def parse_stock(stock_raw: str) -> str:
    """Extract numeric stock from '12 in stock'."""
    stock_raw = clean(stock_raw)
    if not stock_raw:
        return ""
    m = re.search(r"(\d+)", stock_raw)
    return m.group(1) if m else stock_raw


def make_seo_description(desc: str) -> str:
    """Create SEO description from plain text, max 320 chars."""
    text = re.sub(r"<[^>]+>", " ", desc)  # strip HTML
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) > 320:
        text = text[:317] + "..."
    return text


def normalise_tag_value(val: str) -> str:
    """Normalise a value for use in tags."""
    s = val.lower().strip()
    s = re.sub(r"\s+", "-", s)
    s = re.sub(r"[^a-z0-9\-]", "", s)
    s = re.sub(r"-{2,}", "-", s)
    s = s.strip("-")
    return s


def title_case_category(cat: str) -> str:
    """Convert 'ESSENTIALS & ACCESSORIES' → 'Essentials & Accessories'."""
    return " ".join(
        w.capitalize() if w.lower() not in ("&", "and", "or") else w
        for w in cat.split()
    )


def generate_sku_suffix(variant_value: str) -> str:
    """Generate a SKU suffix from a variant value."""
    s = variant_value.upper()
    s = s.replace(" ", "")
    s = s.replace("°", "")
    s = s.replace("\u2033", "IN")  # ″
    s = s.replace('"', "IN")
    s = s.replace("\u201c", "IN").replace("\u201d", "IN")
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s


def build_metadata_comment(record: dict) -> str:
    """Build hidden HTML comment with all WYN source fields."""
    lines = ["<!-- WYN_METADATA"]
    for key in ["Categories_URL", "URL", "Product_URL", "Category", "Category_URL",
                "Categories", "Stock", "RawPrice", "VariantsRaw"]:
        val = record.get(key, "")
        if val:
            # Escape any --> sequences in the value
            val = str(val).replace("-->", "-- >")
            lines.append(f"{key}={val}")
    lines.append("-->")
    return "\n".join(lines)


def empty_row() -> dict:
    """Return a dict with all Shopify headers set to empty string."""
    return {h: "" for h in SHOPIFY_HEADERS}


# ---------------------------------------------------------------------------
# Reference CSV loader
# ---------------------------------------------------------------------------

def load_reference(path: str) -> dict:
    """
    Load the Shopify reference CSV.
    Returns a dict keyed by Variant SKU with the product data.
    """
    ref = {}
    with open(path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        current_handle_data = {}
        for row in reader:
            sku = clean(row.get("Variant SKU", ""))
            handle = clean(row.get("Handle", ""))

            # Track latest product-level data by handle
            if handle and clean(row.get("Title", "")):
                current_handle_data[handle] = {
                    "Handle": handle,
                    "Title": clean(row.get("Title", "")),
                    "Body (HTML)": clean(row.get("Body (HTML)", "")),
                    "Type": clean(row.get("Type", "")),
                    "Tags": clean(row.get("Tags", "")),
                    "SEO Title": clean(row.get("SEO Title", "")),
                    "SEO Description": clean(row.get("SEO Description", "")),
                    "Variant Price": clean(row.get("Variant Price", "")),
                    "Cost per item": clean(row.get("Cost per item", "")),
                }

            if sku:
                entry = current_handle_data.get(handle, {}).copy()
                entry["Variant Price"] = clean(row.get("Variant Price", ""))
                entry["Cost per item"] = clean(row.get("Cost per item", ""))
                ref[sku] = entry

    return ref


# ---------------------------------------------------------------------------
# WYN loader
# ---------------------------------------------------------------------------

def load_wyn(path: str) -> list[dict]:
    """Load WYN xlsx Sheet1 into a list of dicts."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Sheet1"]
    rows = []
    headers = None
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        vals = [clean(v) for v in row]
        if i == 0:
            headers = vals
            continue
        record = {}
        for j, h in enumerate(headers):
            record[h] = vals[j] if j < len(vals) else ""
        rows.append(record)
    wb.close()
    return rows


# ---------------------------------------------------------------------------
# Deduplication
# ---------------------------------------------------------------------------

def deduplicate_wyn(rows: list[dict]) -> list[dict]:
    """
    Deduplicate WYN rows by URL. Merge category and image data.
    """
    groups: dict[str, list[dict]] = OrderedDict()
    for r in rows:
        url = normalise_url(r.get("URL", ""))
        if not url:
            url = normalise_url(r.get("Product_URL", ""))
        if not url:
            continue
        groups.setdefault(url, []).append(r)

    result = []
    for url, group in groups.items():
        primary = group[0].copy()

        if len(group) > 1:
            # Merge categories
            all_cats = set()
            all_cats_url = set()
            all_categories = set()
            all_category_url = set()
            all_images_set = []
            all_images_seen = set()

            for rec in group:
                if rec.get("Categories"):
                    all_cats.add(rec["Categories"])
                if rec.get("Categories_URL"):
                    all_cats_url.add(rec["Categories_URL"])
                if rec.get("Category"):
                    all_categories.add(rec["Category"])
                if rec.get("Category_URL"):
                    all_category_url.add(rec["Category_URL"])

                # Merge images
                imgs = parse_images(rec.get("Main_Image", ""), rec.get("All_Images", ""))
                for img in imgs:
                    if img not in all_images_seen:
                        all_images_seen.add(img)
                        all_images_set.append(img)

                # Keep non-empty SKU
                if not primary.get("SKU") and rec.get("SKU"):
                    primary["SKU"] = rec["SKU"]

                # Keep stock info
                if not primary.get("Stock") and rec.get("Stock"):
                    primary["Stock"] = rec["Stock"]

            primary["Categories"] = ", ".join(sorted(all_cats))
            primary["Categories_URL"] = ", ".join(sorted(all_cats_url))
            primary["Category"] = ", ".join(sorted(all_categories))
            primary["Category_URL"] = ", ".join(sorted(all_category_url))
            primary["_merged_images"] = all_images_set
        else:
            primary["_merged_images"] = parse_images(
                primary.get("Main_Image", ""), primary.get("All_Images", "")
            )

        result.append(primary)

    return result


# ---------------------------------------------------------------------------
# Product building
# ---------------------------------------------------------------------------

def build_products(wyn_rows: list[dict], ref: dict) -> list[dict]:
    """
    Convert deduplicated WYN rows into Shopify product structures.
    Returns a list of product dicts, each containing variant info and images.
    """
    products = []
    used_handles = set()

    for rec in wyn_rows:
        name = rec.get("Name", "")
        if not name:
            continue

        sku = rec.get("SKU", "")
        price_raw = rec.get("Price", "")
        price = parse_price(price_raw)

        # Parse variants
        variant_cols = [rec.get(f"Variant_{i:02d}", "") for i in range(1, 9)]
        option_name, variant_values = parse_variants(
            rec.get("Variants", ""), variant_cols
        )

        # Images
        images = rec.get("_merged_images", [])
        if not images:
            images = parse_images(rec.get("Main_Image", ""), rec.get("All_Images", ""))

        # Determine handle
        ref_match = ref.get(sku) if sku else None
        if ref_match and ref_match.get("Handle"):
            handle = ref_match["Handle"]
        else:
            url = rec.get("URL", "") or rec.get("Product_URL", "")
            handle = url_to_slug(url) if url else slugify(name)

        # Ensure unique handle
        base_handle = handle
        counter = 2
        while handle in used_handles:
            handle = f"{base_handle}-{counter}"
            counter += 1
        used_handles.add(handle)

        # Title
        if ref_match and ref_match.get("Title"):
            title = ref_match["Title"]
        else:
            title = name.upper()

        # Body HTML
        if ref_match and ref_match.get("Body (HTML)"):
            body_html = ref_match["Body (HTML)"]
        elif rec.get("Description_HTML"):
            body_html = rec["Description_HTML"]
        elif rec.get("Description"):
            paragraphs = re.split(r"\n\s*\n", rec["Description"])
            body_html = "".join(f"<p>{html.escape(p.strip())}</p>" for p in paragraphs if p.strip())
            if not body_html:
                body_html = f"<p>{html.escape(rec['Description'])}</p>"
        else:
            body_html = ""

        # Append metadata
        metadata_record = {
            "Categories_URL": rec.get("Categories_URL", ""),
            "URL": rec.get("URL", ""),
            "Product_URL": rec.get("Product_URL", ""),
            "Category": rec.get("Category", ""),
            "Category_URL": rec.get("Category_URL", ""),
            "Categories": rec.get("Categories", ""),
            "Stock": parse_stock(rec.get("Stock", "")),
            "RawPrice": price_raw,
            "VariantsRaw": rec.get("Variants", "").replace("\n", " ").replace("\r", " "),
        }
        metadata = build_metadata_comment(metadata_record)
        if body_html:
            body_html = body_html + "\n" + metadata
        else:
            body_html = metadata

        # Type
        if ref_match and ref_match.get("Type"):
            ptype = ref_match["Type"]
        else:
            cat = rec.get("Category", "")
            # Use first category if multiple
            if "," in cat:
                cat = cat.split(",")[0].strip()
            ptype = title_case_category(cat) if cat else ""

        # Tags
        if ref_match and ref_match.get("Tags"):
            tags = ref_match["Tags"]
        else:
            tag_parts = ["source:wyn"]
            # Category tags
            categories_raw = rec.get("Categories", "")
            if categories_raw:
                for c in categories_raw.split(","):
                    c = c.strip()
                    # Remove "Category: " prefix
                    c = re.sub(r"^Category:\s*", "", c, flags=re.IGNORECASE)
                    if c:
                        tag_parts.append(f"category:{normalise_tag_value(c)}")

            cat_raw = rec.get("Category", "")
            if cat_raw:
                for c in cat_raw.split(","):
                    c = c.strip()
                    if c:
                        tag_parts.append(f"wyn_category:{normalise_tag_value(c)}")

            tags = ", ".join(tag_parts)

        # SEO
        if ref_match and ref_match.get("SEO Title"):
            seo_title = ref_match["SEO Title"]
        else:
            seo_title = title

        if ref_match and ref_match.get("SEO Description"):
            seo_desc = ref_match["SEO Description"]
        else:
            desc_source = rec.get("Description", "") or rec.get("Description_HTML", "")
            seo_desc = make_seo_description(desc_source) if desc_source else ""

        # Build variant rows
        variants = []
        if not variant_values or len(variant_values) <= 1:
            # Single variant product
            opt_name = "Title"
            opt_value = "Default Title"

            if price is not None:
                var_price = format_price(price)
                cost = compute_cost(price)
            elif ref_match and ref_match.get("Variant Price"):
                var_price = ref_match["Variant Price"]
                cost = ""
            else:
                var_price = ""
                cost = ""

            final_sku = sku if sku else handle.upper().replace("-", "")

            variants.append({
                "option_name": opt_name,
                "option_value": opt_value,
                "sku": final_sku,
                "price": var_price,
                "cost": cost,
            })
        else:
            # Multi-variant product
            if not option_name:
                option_name = "Size"

            base_sku = sku if sku else handle.upper().replace("-", "")

            for vval in variant_values:
                suffix = generate_sku_suffix(vval)
                var_sku = f"{base_sku}-{suffix}" if suffix else base_sku

                # Price — use product price for all variants
                if price is not None:
                    var_price = format_price(price)
                    cost = compute_cost(price)
                elif ref_match and ref_match.get("Variant Price"):
                    var_price = ref_match["Variant Price"]
                    cost = ""
                else:
                    var_price = ""
                    cost = ""

                variants.append({
                    "option_name": option_name,
                    "option_value": vval,
                    "sku": var_sku,
                    "price": var_price,
                    "cost": cost,
                })

        products.append({
            "handle": handle,
            "title": title,
            "body_html": body_html,
            "type": ptype,
            "tags": tags,
            "seo_title": seo_title,
            "seo_desc": seo_desc,
            "images": images,
            "variants": variants,
        })

    return products


# ---------------------------------------------------------------------------
# CSV output
# ---------------------------------------------------------------------------

def product_to_rows(product: dict) -> list[dict]:
    """Convert a product dict into a list of Shopify CSV row dicts."""
    rows = []
    handle = product["handle"]
    variants = product["variants"]
    images = product["images"]

    for vi, variant in enumerate(variants):
        row = empty_row()
        row["Handle"] = handle

        if vi == 0:
            # Primary product row
            row["Title"] = product["title"]
            row["Body (HTML)"] = product["body_html"]
            row["Vendor"] = "What You Need"
            row["Type"] = product["type"]
            row["Tags"] = product["tags"]
            row["Published"] = "true"
            row["Gift Card"] = "false"
            row["SEO Title"] = product["seo_title"]
            row["SEO Description"] = product["seo_desc"]
            row["Status"] = "active"

            # First image on primary row
            if images:
                row["Image Src"] = images[0]
                row["Image Position"] = "1"
                row["Image Alt Text"] = product["title"]

        # Variant fields (on every variant row)
        row["Option1 Name"] = variant["option_name"] if vi == 0 else ""
        row["Option1 Value"] = variant["option_value"]
        row["Variant SKU"] = variant["sku"]
        row["Variant Grams"] = "0.0"
        row["Variant Inventory Policy"] = "deny"
        row["Variant Fulfillment Service"] = "manual"
        row["Variant Price"] = variant["price"]
        row["Variant Requires Shipping"] = "true"
        row["Variant Taxable"] = "true"
        row["Variant Weight Unit"] = "lb"
        row["Cost per item"] = variant["cost"]

        rows.append(row)

    # Additional image-only rows (images 2..N)
    if len(images) > 1:
        for img_idx, img_url in enumerate(images[1:], start=2):
            row = empty_row()
            row["Handle"] = handle
            row["Image Src"] = img_url
            row["Image Position"] = str(img_idx)
            rows.append(row)

    return rows


def write_csv(products: list[dict], output_path: str):
    """Write all products to a Shopify-formatted CSV."""
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=SHOPIFY_HEADERS, extrasaction="ignore")
        writer.writeheader()

        for product in products:
            rows = product_to_rows(product)
            for row in rows:
                # Ensure no NaN/None strings
                cleaned = {}
                for k, v in row.items():
                    if v in (None, "None", "NaN", "null"):
                        cleaned[k] = ""
                    else:
                        cleaned[k] = v
                writer.writerow(cleaned)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Transform WYN catalogue into Shopify import CSV"
    )
    parser.add_argument(
        "--wyn",
        default="WYN 31.01.26.xlsx",
        help="Path to WYN catalogue Excel file",
    )
    parser.add_argument(
        "--reference",
        default="products_export_1.csv",
        help="Path to Shopify reference CSV export",
    )
    parser.add_argument(
        "--output",
        default="shopify_import.csv",
        help="Path for output Shopify import CSV",
    )
    args = parser.parse_args()

    print(f"Loading Shopify reference: {args.reference}")
    ref = load_reference(args.reference)
    print(f"  Loaded {len(ref)} existing SKUs from reference")

    print(f"Loading WYN catalogue: {args.wyn}")
    wyn_raw = load_wyn(args.wyn)
    print(f"  Loaded {len(wyn_raw)} raw rows")

    print("Deduplicating WYN rows...")
    wyn_deduped = deduplicate_wyn(wyn_raw)
    print(f"  {len(wyn_deduped)} unique products after deduplication")

    print("Building Shopify products...")
    products = build_products(wyn_deduped, ref)
    print(f"  Built {len(products)} products")

    total_variants = sum(len(p["variants"]) for p in products)
    total_images = sum(len(p["images"]) for p in products)
    print(f"  Total variants: {total_variants}")
    print(f"  Total images: {total_images}")

    print(f"Writing output: {args.output}")
    write_csv(products, args.output)

    # Validate output
    with open(args.output, "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        header = next(reader)
        num_cols = len(header)
        row_count = 0
        errors = []
        for i, row in enumerate(reader, start=2):
            row_count += 1
            if len(row) != num_cols:
                errors.append(f"Row {i}: expected {num_cols} cols, got {len(row)}")
            for j, val in enumerate(row):
                if val in ("NaN", "None", "null"):
                    errors.append(f"Row {i}, col {j} ({header[j]}): invalid value '{val}'")

        if errors:
            print(f"\nValidation errors ({len(errors)}):")
            for e in errors[:20]:
                print(f"  {e}")
        else:
            print(f"\nValidation passed: {row_count} data rows, {num_cols} columns each")

    print("Done!")


if __name__ == "__main__":
    main()
