from bs4 import BeautifulSoup
import requests

def get_the_page_data(pagenumber):
    page = str("https://www.floridabar.org/directories/find-mbr/?lName=&sdx=N&fName=&eligible=Y&deceased=N&firm=&locValue=&locType=C&pracAreas=F01&lawSchool=&services=&langs=&certValue=&pageNumber="+str(pagenumber)+"&pageSize=50")
    html_text = requests.get(page).text
    soup = BeautifulSoup(html_text, 'lxml')
    info = soup.find_all('li', class_ = 'profile-compact')
    for eachinfo in info:
        # print(eachinfo)
        name = eachinfo.find('p',class_ = 'profile-name').text
        # print(name)
        address_number_email = eachinfo.find('div', class_ = 'profile-contact')
        # print(address_number_email) 
        try:
            address = str(address_number_email.find('p'))
            address = address.replace('<br/>','^^')
            address = address.replace('<p>','')
            address = address.replace('</p>','')
            # print(address)
            phone = address_number_email.find('a').text
            # print(phone)

            # the email is encrypted, so here is the decoding process
            emailhtmlstring = str(address_number_email.find('a', class_ = 'icon-email'))
            # print(emailhtmlstring)
            codestring1 = emailhtmlstring.split('#')[1]
            codestring2 = codestring1.split('"')[0]
            # print(codestring2)
            def decodeEmail(e):
                de = "543931142127353935313e352e7a373b39"
                k = int(e[:2], 16)

                for i in range(2, len(e)-1, 2):
                    de += chr(int(e[i:i+2], 16)^k)
                return de
            emailcodestring = decodeEmail(codestring2)
            email = emailcodestring.replace('543931142127353935313e352e7a373b39','')
            # print(email) # finally we get the email
        except AttributeError:
            address = 'NONE'
            phone = 'NONE'
            email = 'NONE'
        except IndexError:
            email = 'NONE'
        ps = [name, address, phone, email] # putting them into a string called ps
        # print(ps)
        with open(f'results2.txt', 'a') as f:
            f.write(f'{ps} \n')
            f.close()

pagenumber = 56
for i in range(56,63):
    pagenumber +=1
    print(pagenumber)
    get_the_page_data(pagenumber)
    print(f'finish page {pagenumber}')