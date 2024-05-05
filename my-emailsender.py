from emailsender import *

#Start Process Timer
start = time.time()    

path = file_path()
emails_list = pd.read_excel(path, dtype=str)
SUBJECT = "Introducing Sterk Engineering"

i = int(input('Start Line\n(Zero for First Line): '))
n = int(input('End Line\n(Will be included): '))
for i in range(i, n+1):
        receiver_info = get_receiver(i, emails_list)
        print(receiver_info)

        content = read_message()
        content = content.replace('Salutation', receiver_info['salutation'])
        send_email(SUBJECT, receiver_info['to'], content, receiver_info['date'])

else:
     print(f"emails_list between {i} and {n+1} are sent!")

#End Process Timer
end = time.time()
#Time elapsed
print(end-start)