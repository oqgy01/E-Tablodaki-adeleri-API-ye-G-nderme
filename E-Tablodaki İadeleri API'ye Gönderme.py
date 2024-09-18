import requests
import pandas as pd
import xml.etree.ElementTree as ET
import hashlib
import codecs



# İlk API fonksiyonları
def create_xml_request(username, password, client_id, type, order_id, total=None):
    root = ET.Element("CC5Request")
    ET.SubElement(root, "Name").text = username
    ET.SubElement(root, "Password").text = password
    ET.SubElement(root, "ClientId").text = client_id
    ET.SubElement(root, "Type").text = type
    ET.SubElement(root, "OrderId").text = order_id
    if total is not None:
        ET.SubElement(root, "Total").text = str(total)
    return ET.tostring(root, encoding='utf-8', method='xml').decode()

def send_request(url, xml_data):
    headers = {'Content-Type': 'application/xml'}
    response = requests.post(url, data=xml_data, headers=headers)
    return response

# İkinci API fonksiyonları
def sha1(text):
    text_bytes = codecs.encode(text, 'iso-8859-9')
    sha1_hash = hashlib.sha1(text_bytes).hexdigest().upper()
    return sha1_hash

def get_hash_data(user_password, terminal_id, order_id, amount, currency_code):
    hashed_password = sha1(user_password + "0" + str(terminal_id))
    hash_data = sha1(str(order_id) + str(terminal_id) + str(amount) + hashed_password)
    return hash_data

def create_gvps_request(mode, version, prov_user_id, hash_data, user_id, terminal_id, merchant_id, ip_address, email, order_id, amount, currency_code, transaction_type="refund"):
    root = ET.Element("GVPSRequest")
    ET.SubElement(root, "Mode").text = mode
    ET.SubElement(root, "Version").text = version
    terminal = ET.SubElement(root, "Terminal")
    ET.SubElement(terminal, "ProvUserID").text = prov_user_id
    ET.SubElement(terminal, "HashData").text = hash_data
    ET.SubElement(terminal, "UserID").text = user_id
    ET.SubElement(terminal, "ID").text = str(terminal_id)
    ET.SubElement(terminal, "MerchantID").text = merchant_id
    customer = ET.SubElement(root, "Customer")
    ET.SubElement(customer, "IPAddress").text = ip_address
    ET.SubElement(customer, "EmailAddress").text = email
    order = ET.SubElement(root, "Order")
    ET.SubElement(order, "OrderID").text = str(order_id)
    transaction = ET.SubElement(root, "Transaction")
    ET.SubElement(transaction, "Type").text = transaction_type
    ET.SubElement(transaction, "Amount").text = str(amount)
    ET.SubElement(transaction, "CurrencyCode").text = currency_code
    ET.SubElement(transaction, "CardholderPresentCode").text = "0"
    ET.SubElement(transaction, "MotoInd").text = "H"
    return ET.tostring(root, encoding='utf-8', method='xml').decode()

# API parametreleri
username = "haydigiyadmin"
password = "HYGD4665"
client_id = "700677754665"
api_url_isbank = "https://vpos3.isbank.com.tr/fim/api"

user_password = "Mustafa.51"
terminal_id = "10233352"
currency_code = "949"
prov_user_id = "PROVRFN"
user_id = "PROVAUT"
merchant_id = "1799961"
ip_address = "192.168.0.1"
email = "info@haydigiy.com"
api_url_garanti = "https://sanalposprov.garanti.com.tr/VPServlet"

# Excel dosyasını okuma
df = pd.read_excel('ent-iadeler.xlsx', engine='openpyxl')

# Sonuçları işleme
results = []

for _, row in df.iterrows():
    order_id = row['SiparisId']
    total = row['IBAN']
    bank_name = row['Banka Adı']
    
    if bank_name == "İş Bankası":
        xml_data = create_xml_request(username, password, client_id, "Credit", str(order_id), str(total))
        response = send_request(api_url_isbank, xml_data)
        response_text = response.text
        root = ET.fromstring(response_text)
        error_msg = root.findtext('.//ErrMsg')
        if error_msg:
            results.append((order_id, total, 'hatalı', error_msg))
        else:
            results.append((order_id, total, 'başarılı', None))
    
    elif bank_name == "Garanti":
        amount = str(total)  # Total değeri olduğu gibi kullan
        hash_data = get_hash_data(user_password, terminal_id, str(order_id), amount, currency_code)
        xml_data = create_gvps_request("PROD", "1", prov_user_id, hash_data, user_id, str(terminal_id), merchant_id, ip_address, email, str(order_id), amount, currency_code)
        response = send_request(api_url_garanti, xml_data)
        response_text = response.text
        root = ET.fromstring(response_text)
        error_msg = root.findtext('.//ErrorMsg')
        if error_msg:
            results.append((order_id, total, 'hatalı', error_msg))
        else:
            results.append((order_id, total, 'başarılı', None))
    else:
        results.append((order_id, total, 'banka adı geçersiz', None))

# Sonuçları Excel dosyasına ekleme
results_df = pd.DataFrame(results, columns=['SiparisId', 'IBAN', 'Durum', 'Hata Mesajı'])
df = df.merge(results_df, on=['SiparisId', 'IBAN'], how='left')
df.to_excel('ent-iadeler.xlsx', index=False, engine='openpyxl')
