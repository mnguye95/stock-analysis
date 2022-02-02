from urllib.request import urlopen
import requests
import random

# Credit: https://github.com/clarketm/proxy-list
proxy_list_url = "https://raw.githubusercontent.com/clarketm/proxy-list/master/proxy-list-raw.txt"

ip_addresses = urlopen(proxy_list_url)
ip_addresses = list(map(lambda x: x.decode("utf-8").replace('\n', ''), ip_addresses))

def grab_proxy():
    while 1:
        try:
            proxy_index = random.randint(0, len(ip_addresses) - 1)
            current_ip = ip_addresses[proxy_index]
            proxies = {
                "http": "http://{}".format(current_ip), 
                "https": "https://{}".format(current_ip)
            }
            requests.get("https://www.google.com/", proxies=proxies, timeout=3)
            return current_ip
        except Exception as e:
            ip_addresses.pop(proxy_index)
    return None