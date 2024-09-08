from functions import *

#menu
while True:
    print("\n1. add client\n2. search client\n3. exit")
    menu_id = input("Podaj wartość: ")
    if menu_id == "1":
        add_new_buyer()
    if menu_id == "2":
        list_of_buyers()
    if menu_id == "3":
        break
'''
set_data_of_invoice()
insert_seller_details()
insert_buyer_details()
add_item_to_invoice(1)


n = 1
while True:
    print("\n x - wyjście \n 1 - dodaj następną pozycję")
    your_choice = input("Podaj wartość: ")
    if your_choice != "x":
        n += 1
        add_item_to_invoice(n)

    else:
        break

insert_total_amount()
payment_information()
save_as_docx()
convert_to_pdf()
move_file()
'''
