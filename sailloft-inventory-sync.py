import dropbox
import pandas as pd
import requests
from io import BytesIO

# ================================
# KONFIGURATION
# ================================

# Dropbox-Konfiguration
DROPBOX_ACCESS_TOKEN = "sl.u.AFo8rntdkQk9iedd7X0KtDIGtBuJZChBVKxtJfb2ags6THZRqBpHhfOuSBMKfXuY9XeQA6XOn_ZVJv5e28afEW8ecFz5eHiSlXfwbULiCrZyfIDF2McpiRiV_w54SdSuGHp5HvUDXTjvSe0of7XhmzLZwF98JoPEYB2tD_fqdIN3OoXRfrC3xmL4wbEXIZXHivuvW4zhY9fRyM6kcnLI6BKKCpCGtXWrPQmHLHYVP4au0L6x-FKablsk2oVchXduiDAxP6sV1lQgMs90jk0msJUAuYWUpkqn28I-rr99u0680l2sb6wTL2uflXSwsPT-G9whLjmRP2vxqjasLIWk86LOkTFQbtdf3lDWyXjyNb00cO5NIPQ8Kyhae7LOrslMZV-up72W3MnBOv4ZehRhaoMKgQBNSpoZhmKid3OacXDN1Eol0kmSP2B-zLL82u_fwUoFzBehbe6Kf9wIiJ7Cyo-22LB2klty-ULlkgVO4rLftOpfQxK1tl_coSjLd6ZAQiYF7fc7x6ilsnUnHESFt6RXeZBgLdMEVmI8-NvJNtKlIkLqjh7XAsoXH9bxXESEnd7xtsE3EKgcRBaP8IOJPJhMIWYA5utBPDOP-k-8RmS_a06oyqOoaFSfErcOQAFGiWq0xE48omTUL0kQ94VEEOlY1AJU-pxB3CdkZiAIB9ljAjIScVXeh8lmZQCLvRzUQQrE4gr1OjS3GEp07DDEqESGvFXxrmMgvVxGQl6bsfq9EAwg28tA3awLdE8EMjsZXJ1KLDnuxIkWUBfoy0zeWge-Lp1A-XKbXQahNs4dUx23k18q2uDJd3AfrzHjqi5jTPPOWvutHHvTC9X9hRINTiQSBXhLa2AD4nMJ6YIMv1sqFpSfY1xKtXHljYGX9D7GGq4WBRpX8EIh_p8ww-17-gDystIRO7Qs4RbwzUVSNccrzLpUbBpFOt_gXWerPCyRHdiU33kdwC9Tr-Bzi-5-d1rCsBb20X4V9MzuYT6-ZBkDrErAYQWrvvCayRakqsIAZokNOrYS8hj8Hi7x-VFUwUlaD_rnuEMsbAp8qMKn9c4WpybgGrchAZVYMgbML-5byafJqtCRHKPS7ADqTI6ZH7dRZWROA6ZCpXcNdlk9awtZ-AqSDVyypXiCI7EwtA1ZwR6Ia-tBjE2Dw0hLqapIVlAIy_GmTUHLULKHM377R-qage0icikPsCmRfEtQ1HLG8YJ3cjjAilh8g187JQ4nNEdEMlkdA1zMjJzP7jcOR16bUmRmA1bG8y1IjukfHC6T1Aw"  # Ersetze diesen Platzhalter
DROPBOX_FILE_PATH = "/Sailloft_Masts_2025_Stock_.xlsx"  # Pfad in deiner Dropbox

# Shopify-Konfiguration
# Wichtig: Für API-Anfragen muss die myshopify-Domain verwendet werden (nicht die Storefront-Domain)
SHOP_URL = "https://xrivy0-jx.myshopify.com"  # Ersetze diesen Platzhalter, z. B. "https://shop-sailloft-de.myshopify.com"
API_VERSION = "2024-01"
SHOPIFY_ACCESS_TOKEN = "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"  # Ersetze diesen Platzhalter
LOCATION_ID = "108001198419"  # Die bereits ermittelte Lagerort-ID

# ================================
# FUNKTIONEN
# ================================

def download_excel_from_dropbox():
    """
    Lädt die Excel-Datei aus Dropbox herunter und gibt einen BytesIO-Stream zurück.
    """
    dbx = dropbox.Dropbox(DROPBOX_ACCESS_TOKEN)
    metadata, res = dbx.files_download(DROPBOX_FILE_PATH)
    return BytesIO(res.content)

def load_inventory_data(file_bytes):
    """
    Liest die Lagerdaten aus der Excel-Datei.
    Annahme: Das relevante Arbeitsblatt heißt "Sailoft 2025" 
    und die Daten beginnen nach 10 übersprungenen Zeilen.
    """
    df = pd.read_excel(file_bytes, sheet_name="Sailoft 2025", skiprows=10, engine="openpyxl")
    # Entferne Zeilen ohne gültige SKU oder Stock-Wert
    df = df.dropna(subset=["SKU", "Stock"])
    # Falls sich wiederholt Header-Zeilen eingeschlichen haben, überspringe diese:
    df = df[df["SKU"] != "SKU"]
    return df

def get_inventory_item_id_by_sku(sku):
    """
    Nutzt Shopify GraphQL, um anhand der SKU die inventory_item_id zu ermitteln.
    Gibt die ID (nur die numerische Komponente) als String zurück oder None, wenn nicht gefunden.
    """
    query = {
        "query": f"""
        {{
            productVariants(first: 1, query: "sku:{sku}") {{
                edges {{
                    node {{
                        inventoryItem {{
                            id
                        }}
                    }}
                }}
            }}
        }}
        """
    }
    headers = {
        "X-Shopify-Access-Token": SHOPIFY_ACCESS_TOKEN,
        "Content-Type": "application/json"
    }
    url = f"{SHOP_URL}/admin/api/{API_VERSION}/graphql.json"
    response = requests.post(url, json=query, headers=headers)
    if response.status_code != 200:
        print(f"GraphQL error for SKU {sku}: {response.status_code} {response.text}")
        return None
    data = response.json()
    try:
        # Beispiel: "gid://shopify/InventoryItem/1234567890"
        inventory_item_id = data["data"]["productVariants"]["edges"][0]["node"]["inventoryItem"]["id"]
        return inventory_item_id.split("/")[-1]
    except (KeyError, IndexError):
        print(f"SKU {sku} nicht in Shopify gefunden.")
        return None

def update_inventory_level(inventory_item_id, quantity):
    """
    Aktualisiert über die Shopify REST API den verfügbaren Lagerbestand 
    für einen gegebenen inventory_item_id und den definierten LOCATION_ID.
    """
    url = f"{SHOP_URL}/admin/api/{API_VERSION}/inventory_levels/set.json"
    headers = {
        "X-Shopify-Access-Token": SHOPIFY_ACCESS_TOKEN,
        "Content-Type": "application/json"
    }
    payload = {
        "location_id": LOCATION_ID,
        "inventory_item_id": inventory_item_id,
        "available": quantity
    }
    response = requests.post(url, headers=headers, json=payload)
    if response.status_code == 200:
        print(f"✅ Bestand für Inventory Item {inventory_item_id} auf {quantity} gesetzt.")
    else:
        print(f"❌ Fehler beim Updaten von Inventory Item {inventory_item_id}: {response.status_code}, {response.text}")

def sync_inventory():
    """
    Hauptfunktion zur Synchronisierung der Lagerdaten:
      - Lädt die Excel-Datei aus Dropbox
      - Liest die Daten ein
      - Aktualisiert für jede gültige SKU den Bestand in Shopify
    """
    try:
        file_bytes = download_excel_from_dropbox()
    except Exception as e:
        print(f"Fehler beim Herunterladen der Dropbox-Datei: {e}")
        return
    
    try:
        df = load_inventory_data(file_bytes)
    except Exception as e:
        print(f"Fehler beim Lesen der Excel-Datei: {e}")
        return
    
    print("Starte Lagerbestand-Synchronisierung...")
    for index, row in df.iterrows():
        sku = str(row["SKU"]).strip()
        try:
            quantity = int(row["Stock"])
        except Exception as e:
            print(f"Ungültiger Stock-Wert für SKU {sku}: {row['Stock']}")
            continue
        
        print(f"Verarbeite SKU: {sku} mit Bestand: {quantity}")
        inventory_item_id = get_inventory_item_id_by_sku(sku)
        if inventory_item_id:
            update_inventory_level(inventory_item_id, quantity)
        else:
            print(f"Kann Inventory Item für SKU {sku} nicht ermitteln.")

# ================================
# MAIN
# ================================

if __name__ == "__main__":
    sync_inventory()
