import speech_recognition as sr
import pyttsx3
import os
import boto
import re
import urllib
import webbrowser
import smtplib
import requests
import subprocess
from pyowm import OWM
from time import strftime
import json
import psutil
import time
import WMI
import wikipedia

from bs4 import BeautifulSoup as BeautifulSoup
import random
import pyautogui

speech= sr.Recognizer()
greeting_dict={'hello':'hello','hi':'hi'}
mp3_greeting_dict={'Yes mr. Sunny','Yes mr. Sunny'}
open_launch_dict={'open':'open','launch':'launch'}
social_media_dict={'facebook':'http://www.facebook.com','twitter':'http://www.twitter.com'}

try:
    engine = pyttsx3.init()
except ImportError:
    print('Request')
except RuntimeError:
    pass

voices = engine.getProperty('voices')
for voice in voices:
    print(voice.id)
engine.setProperty('voice','HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Tokens\TTS_MS_EN-US_ZIRA_11.0')
rate= engine.getProperty('rate')
engine.setProperty('rate',rate)

def speak_text_cmd(cmd):
    engine.say(cmd)
    engine.runAndWait()

def read_voice_cmd():
    with sr.Microphone() as source:
        audio=speech.listen(source=source,timeout=10,phrase_time_limit=5)
    try:
        voice_text=speech.recognize_google(audio)
    except sr.RequestError:
        print("Network Error")
    except sr.WaitTimeoutError:
        pass
    return voice_note
    
    

def is_valid_dict(greeting_dict,voice_note):
    
    for key, value in greeting_dict.item():
        try:
            if value==voice_note.split('')[0]:
                return True
    
            elif key== voice_note.split('')[1]:
                return True
        except IndexError:
            pass
    return False

def secs2hours(secs):
    mm,ss=divmod(secs, 60)
    hh,mm=divmod(mm, 60)
    return "%dhour, %02d minute, %02s second"%(hh, mm, ss)
if __name__ == "__main__":
    speak_text_cmd('Hello Mr. Sunny this is Alex')
    
    while True:
        voice_note= read_voice_cmd().lower()
        print('cmd : {}'.format(voice_note))
        if is_valid_dict(greeting_dict,voice_note):
            speak_text_cmd(mp3_greeting_dict)
            continue
        elif is_valid_dict(open_launch_dict,voice_note):
            speak_text_cmd("sure sir. ")
            if(is_valid_dict(social_media_dict,voice_note)):
                key=voice_note.split(' ')[1]
                webbrowser.open(social_media_dict.get(key))
        elif 'open' in voice_note:
            reg_ex= re.search('open (.*)', voice_note)
            if reg_ex:
                domain = reg_ex.group(1)
                print(domain)
                url = 'https://www.' + domain
                webbrowser.open(url)
                speak_text_cmd('The website you have requested has been opened for you  Sir. ')
            continue
        elif is_valid_dict(open_launch_dict,voice_note):

            webbrowser.open('http://www.google.co.in/search?q={}'.format(voice_note))

        elif 'current weather' in voice_note:
            reg_ex = re.search('current weather in (.*)',voice_note)
            if reg_ex:
                city= reg_ex.group(1)
                owm = OWM(API_key='ab0d5e80c8dafb2cb8lfa9e824lfa')
                obs=owm.weather_at_place(city)
                w=obs.get_weather()
                k=w.get_status()
                x=w.get_temperature(unit = 'celsius')
                print('Current weather in %s is %s. The maximum temperature  is %0.2f and the maximum temperature is %0.2f degree celcius') 
                speak_text_cmd('Current weather in %s is %s. The maximum temperature  is %0.2f and the maximum temperature is %0.2f degree celcius')
        elif 'shut down' in voice_note:
            speak_text_cmd('ok sir')
            speak_text_cmd('shutting down your oprating  system sir. ')
            os.system('shutdown -s')

        elif 'restart' in voice_note:
            speak_text_cmd('ok sir')
            speak_text_cmd('restart your oprating  system sir. ')
            os.system('restart -r')
        elif 'send mail' in voice_note:
            speak_text_cmd('ok sir')
            server = smtplib.SMTP('smtp.gmail.com',587)
            server.ehlo()
            server.starttls()
            server.login('sunnychourasiya541@gmail.com','sunny@chourasiya7440')
            speak_text_cmd('who is the recepient')
            print('tell')
            send=read_voice_cmd().lower()
            speak_text_cmd('what is the subject line')
            print('tell')
            subject= read_voice_cmd().lower()
            print('subject')
            speak_text_cmd('what is the message')
            print('tell')
            message=read_voice_cmd().lower()
            print('message')
            server.sendmail('sunnychourasiya541@gmail.com',send,'Subject: ' +subject+ '\n\n' + message )
        elif 'time' in voice_note:
            currrent_time=time.strftime("%d:%B:%Y:%A:%H:%M:%S")
            print(currrent_time)
            speak_text_cmd('sir, today date is' + time.strftime("%d:%B:%Y"))
            speak_text_cmd(time.strftime("%A"))
            speak_text_cmd('and time is' +time.strftime("%d:%H:%M:%S"))
        elif 'brightness' in voice_note:
            if 'decrease' in voice_note:
                print('ok listen')
                speak_text_cmd('got it sir')
                dec=wmi.WMI(namespace='wmi')
                methods=dec.WmiMonitorBrightnessMethods()[0]
                methods.WmiSetBrightness(30,0)
            elif 'increase' in voice_note:
                speak_text_cmd('got it sir')
                ins=wmi.WMI(namespace='wmi')
                methods=ins.WmiMonitorBrightnessMethods()[0]
                methods.WmiSetBrightness(100,0)
                speak_text_cmd('Brightness increase')
        elif 'charge' in voice_note:
            speak_text_cmd('got it sir')
            battery=psutil.sensors_battery()
            plugged=battery.power_plugged
            percent=int(battery.percent)
            time_left=secs2hours(battery.secsleft)
            print(percent)
            if percent < 40 and plugged== False:
                speak_text_cmd('sir, please connect the charger because i can servive only' +time_left)
            if percent < 40 and plugged== True:
                speak_text_cmd("don't worry,sir charger is connected")
            else:
                speak_text_cmd(' sir,no need  to connect the charger because  i can survive' +time_left)
        elif 'screenshot' in voice_note or 'screen shot' in voice_note or 'snapshot' in voice_note:
            speak_text_cmd('ok sir let me take a snapshot')
            speak_text_cmd("ok done")
            speak_text_cmd('check your Desktop, i saved there')
            pic=pyautogui.screenshot()
            pic.save('C:Users.Hp/Desktop/Screenshot.png')
        elif 'drive' in voice_note:
            speak_text_cmd('opening drive sir')
            drive= voice_note[5]
            os.system('explorer'+drive+'://'.format)
        elif 'bya' in voice_note:
            speak_text_cmd("By Mr. Sunny heppy to help you. Have a good day")
            exit()
            