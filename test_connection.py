import requests
import sys

url = "http://3.142.164.128:8000/tasks/poll"
username = "admin"
password = "admin"

print(f"Testing connection to {url}...")
try:
    resp = requests.get(url, params={"username": username}, auth=(username, password), timeout=5)
    print(f"Status Code: {resp.status_code}")
    print(f"Response: {resp.text}")
except Exception as e:
    print(f"Error: {e}")
