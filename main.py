# -*- coding: UTF-8 -*-
import requests as req
import json,sys,time
#先注册azure应用,确保应用有以下权限:
#files:	Files.Read.All、Files.ReadWrite.All、Sites.Read.All、Sites.ReadWrite.All
#user:	User.Read.All、User.ReadWrite.All、Directory.Read.All、Directory.ReadWrite.All
#mail:  Mail.Read、Mail.ReadWrite、MailboxSettings.Read、MailboxSettings.ReadWrite
#注册后一定要再点代表xxx授予管理员同意,否则outlook api无法调用






path=sys.path[0]+r'/Secret.txt'
num1 = 0

def gettoken(refresh_token):
    headers={'Content-Type':'application/x-www-form-urlencoded'
            }
    data={'grant_type': 'refresh_token',
          'refresh_token': refresh_token,
          'client_id':id,
          'client_secret':secret,
          'redirect_uri':'http://localhost:53682/'
         }
    
    data={'grant_type': 'refresh_token',
          'refresh_token': "0.AX0Ah-A6iYqzLU6HwimXSjrKDLE_H8qCNEJHq_2-lqVc4VWaAMc.AgABAAEAAAD--DLA3VO7QrddgJg7WevrAgDs_wQA9P_GKoED7vZ14dMqEgs4urqmUcHLIJ06wAdMpvBfs0tiGRKK4n5kcXMk1-6vujzZ6pSE4kt2yEr7bI8VfVbEMi5F7Xrq4zlJoYmyRvRmEm_1zWa7i8Gpx73BFltsZkTp1zpjXRm2gKb4eMr_-PG59GerJryUEQAtl7GJgcbJlXgeUltROvRhWcNpUMpB6yiVa0bqnJiJY9yihWBkSvyfLZYtvnHeEl74brIcurROxr0WIdqtAk3Z7Ms80js4o_HaKn2wiNH1ToaqpY5wtyNjhKw05VpRm2Mp8VVJ9cuBdQf-zyHVzF2JRaU6WZtbniAw5tUqZXTkDgo_1z9hhXaAwAgGNU9YWK-1qWzyTppdbjdbEdBVaxV5xggAt8dzxpgQMk0S0g9-GheVgestu667IOtTbTKKuhEEDGadueKbRi3WCki_KbBEi14AZEIOulajR37w7SYlG004Q065yLksBuEtMojjOxT3-6BCf1Af5mENC0ftE3-R8LHaVnzb9XHkyTZCYtNGDlnvW6nhOeZ9Zp2ejgrWR9tBCLO_2kk_lYg8KhgqJDMtrc9H3BRYAH5ZnCGbAyvm_Ho6Anw89K9E_d9cVy_A0Nvbx1fAZs8vHEGG23Fq1t52dl4tb4f71FdHbiX8Z_dJskHHyAEGYb2vpz9Uw1yz-jFY4DW5dhJpy0Uk9sZ9J-pVjHkDqvBhRWDJ0nZeL8BEWNMqLVdAbGtCetcnpHZesVtIzhoXc5L8oNsllxj2N_JlN_40kacwjYcSZ3GYtiv-cAjvRB-G7qdpLa5cOzVk6PuaYdlKvyHk",
          'client_id':"ca1f3fb1-3482-4742-abfd-be96a55ce155",
          'client_secret':"Wvh8Q~BHeYHyVVgE1dPW3fACpuKPGI.jWajabdes",
          'redirect_uri':'http://localhost:53682/'
         }
    
    html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=headers)
    jsontxt = json.loads(html.text)
    refresh_token = jsontxt['refresh_token']
    access_token = jsontxt['access_token']
    with open(path, 'w+') as f:
        f.write(refresh_token)
    return access_token
def main():
    fo = open(path, "r+")
    refresh_token = fo.read()
    fo.close()
    global num1
    localtime = time.asctime( time.localtime(time.time()) )
    access_token=gettoken(refresh_token)
    headers={
    'Authorization':access_token,
    'Content-Type':'application/json'
    }
    try:
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive',headers=headers).status_code == 200:
            num1+=1
            print("2调用成功"+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/drive/root',headers=headers).status_code == 200:
            num1+=1
            print('3调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/users ',headers=headers).status_code == 200:
            num1+=1
            print('4调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/messages',headers=headers).status_code == 200:
            num1+=1
            print('5调用成功'+str(num1)+'次')    
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',headers=headers).status_code == 200:
            num1+=1
            print('6调用成功'+str(num1)+'次')    
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',headers=headers).status_code == 200:
            num1+=1
            print('7调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root/children',headers=headers).status_code == 200:
            num1+=1
            print('8调用成功'+str(num1)+'次')
        if req.get(r'https://api.powerbi.com/v1.0/myorg/apps',headers=headers).status_code == 200:
            num1+=1
            print('8调用成功'+str(num1)+'次') 
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders',headers=headers).status_code == 200:
            num1+=1
            print('9调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/outlook/masterCategories',headers=headers).status_code == 200:
            num1+=1
            print('10调用成功'+str(num1)+'次')
            print('此次运行结束时间为 :', localtime)
    except:
        print("pass")
        pass
for _ in range(3):
    main()
