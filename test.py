import requests

url = "http://10.42.78.181:8501"
x = requests.get(url)
print(x.content)