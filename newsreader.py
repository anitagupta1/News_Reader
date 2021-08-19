import os
import requests
import json
from datetime import datetime
from win32com.client import Dispatch

def speak(s):
    speak=Dispatch('SAPI.SpVoice')
    speak.Speak(s)

newzapi=os.environ['Newzapi']                                               # to extract ur api key from environment variable of ur pc
print("COUNTARY :\n at, au, be, bg, br, ca, ch, cn, co, cu, cz, de, eg, fr, gb, gr \n hk, hu, id, ie, il, in, it, jp, kr, lt, lv, ma, mx, my, ng, nl, nz, ro, rs, \n ru, sa, se, sg, si, sk, th, tr, tw, ua, us, ve, za \n Eg in(india),us(united state),it(italy), etc")
countary_newz=input("select the abbreviated country name from above to listen that country newz \n       ")
print("NEWZ CATEGORIES  :\n business ******* entertainment ******* general ******* health ******* science ******* sports ******* technology \n")
newz_type=input("Select the category of the newz  \n      ")

#try:
r = requests.get(f"https://newsapi.org/v2/top-headlines?country={countary_newz}&category={newz_type}&apiKey={newzapi}")
#finally:
 # r = requests.get(f"https://newsapi.org/v2/top-headlines?country=in&apiKey={newzapi}")

if __name__ == '__main__':
    t=r.text                                                                 #to copy all text from apikey
    q=json.loads(t)                                                          #to convert string into dict
    a=1
    if (q["status"]!="ok"):
        speak("access denied(invalid api key or url)")
    else:
      if (q['totalResults']== 0):
        speak("no newz available ( or SPELL THE COUNTRY AND CATEGORY NAME INCORRECTLY)")
      else:
        n = datetime.now()
        t=n.strftime("%d %B %Y")
        print(n.strftime("%d %B %Y"))
        speak(f"todays date is {t} ")
        print(q['totalResults'])
        for i in range(len(q['articles'])):                                 #to access all the news present in article dictionary
              if (q['articles'][i]['description']) is not None:             #to skip none value of description key
                speak(f' todays newz number {a}')
                print(f" {a} >>>> {q['articles'][i]['description']}")
                speak(q['articles'][i]['description'])
                a+=1

        speak("thanks for listining the newz")

