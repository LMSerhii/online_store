from config import PromExportConfig
from clients import PromClient
import time

def get_all_order_ids(client: PromClient, page) -> list[int]:
    orders = client.get_orders(page=page)

    order_ids = [order["id"] for order in orders if "id" in order ]
    return order_ids


def set_status(client: PromClient, order_ids: list[int]):
    resp = client.set_status(
        order_ids=order_ids,
        new_status_id=3,
        new_status_type="default",
    )

    return resp

def main():
    export_config = PromExportConfig.from_env()
    client = PromClient(export_config)

    page = 1

    while True:
        print("page", page)
        orders_ids = get_all_order_ids(client, page)

        if not orders_ids:
            break

        resp = set_status(client, orders_ids)
        page += 1


        print("resp", resp)
        print("orders_ids_count", len(orders_ids))

        time.sleep(4)








if __name__ == "__main__":
    main()
