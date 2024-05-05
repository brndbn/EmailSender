
from emailsender import *

emails_list = pd.read_excel('email_to_send.xlsx', dtype=str)
SUBJECT = "Introducing Sterk Engineering"

i = int(input('Start Line\n(Zero for First Line): '))
n = int(input('End Line\n(Will be included): '))
for i in range(i, n+1):
        receiver_info = get_receiver(i, emails_list)
        print(type(receiver_info))

        content = read_message()
        content = content.replace('Salutation', receiver_info['salutation'])
        print(content)

else:
     print(f"emails_list between {i} and {n+1} are sent!")