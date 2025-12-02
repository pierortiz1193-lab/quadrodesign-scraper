import requests
import pandas as pd


def main():
    print("Descargando colecciones...")
    collections_url = "https://quadro.work-mgmt.it/api/public/collections"
    collections = requests.get(collections_url).json()

    print("Descargando productos...")
    products_url = "https://quadro.work-mgmt.it/api/public/products"
    products = requests.get(products_url).json()

    # Crear mapa ID->nombre colección
    collection_map = {c["id"]: c["name"] for c in collections}

    rows = []

    print("Agrupando códigos por colección...")
    for prod in products:
        col_id = prod.get("collection_id")
        if not col_id:
            continue

        rows.append({
            "collection": collection_map.get(col_id, "Unknown"),
            "code": prod.get("code", ""),
            "product_name": prod.get("name", "")
        })

    df = pd.DataFrame(rows)
    df = df.sort_values(["collection", "code"])

    print("Guardando Excel...")
    df.to_excel("quadrodesign_codes.xlsx", index=False)

    print("\n✔ Archivo creado: quadrodesign_codes.xlsx")


if __name__ == "__main__":
    main()
