import requests
apiKey = 'ZlA0wiVWdf8AvPodV4MCrXIxs3rcIDAx'
filePath = 'C:\OCR\Pic\label05-03-724x1024.png'
url = "https://api.iapp.co.th/ocr/v3/receipt/file"

headers = {'apikey': apiKey}
files = {'file': ('receipt.jpg', open(filePath, 'rb'), 'image/png')}
data = {'return_image': 'false', 'return_ocr': 'false'}

response = requests.post(url, headers=headers, files=files, data=data)
print(response.json())