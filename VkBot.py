from pyqiwip2p import QiwiP2P
import vk_api
import time
import random
import pandas as pd
import openpyxl

token = ""
p2p = QiwiP2P(auth_key="")


vk = vk_api.VkApi(token=token)
vk._auth_token()

kol_strok = pd.read_excel("games.xlsx")

while True:
    try:
        messages = vk.method("messages.getConversations", {"offset": 0, "count": 20, "filter": "unanswered"})
        if messages["count"] >= 1:
            id = messages["items"][0]["last_message"]["from_id"]
            body = messages["items"][0]["last_message"]["text"]
            if body.lower() == "case1":
                bill = p2p.bill(amount=1, lifetime=10, comment="Абоба")
                vk.method("messages.send", {"peer_id": id, "message": f'Ссылка:\n{bill.pay_url}', "random_id": random.randint(1, 2147483647)})
                while True:
                    if p2p.check(bill_id=bill.bill_id).status == "PAID":
                        x = len(kol_strok)
                        kluch = random.randint(1, x + 1)
                        spisok = openpyxl.open("games.xlsx", read_only=True)
                        sheet = spisok.active
                        kluch_vadacha = sheet[kluch][0].value
                        filename = "games.xlsx"
                        wb = openpyxl.load_workbook(filename)
                        sheet = wb["Лист1"]
                        sheet.delete_rows([kluch][0], 1)
                        wb.save(filename)
                        vk.method("messages.send", {"peer_id": id, "message": f"Оплата успешно прошла!/n Ваш ключ {kluch_vadacha}",
                                                    "random_id": random.randint(1, 2147483647)})
                        break
    except Exception as E:
        time.sleep(1)