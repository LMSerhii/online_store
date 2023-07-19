import time
from tqdm import tqdm
import requests
import lxml

from bs4 import BeautifulSoup
from fake_useragent import UserAgent

ua = UserAgent()


def request_test(url, method, data=None, headers=None, retry=5):
    """ """
    try:
        if method != 'post':
            response = requests.get(url=url, headers=headers)
            # print(f"[+] {url} {response.status_code}")
        else:
            response = requests.post(url=url, data=data,  headers=headers)
            # print(f"[+] {url} {response.status_code}")
    except Exception:
        time.sleep(3)
        if retry:
            # print(f"[INFO] retry={retry} => {url}")
            # tqdm.write(f"[INFO] retry={retry} => {url}")
            return request_test(url, method, data, headers, retry=(retry - 1))
        else:
            raise
    else:
        return response


def main():

    headers = {
        "accept": "*/*",
        "user-agent": ua.random
    }

    example_links = ['']

    for url in example_links:

        try:
            res = request_test(url=url, headers=headers)
            soup = BeautifulSoup(res.text, "lxml")
            print(f"{soup.title.text}\n{'#' * 50}")

        except Exception as ex:
            continue


if __name__ == "__main__":
    main()
