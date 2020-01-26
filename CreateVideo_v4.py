import subprocess
import os
import google.cloud
from google.cloud import texttospeech
import comtypes.client
from pdf2image import convert_from_path, convert_from_bytes
from pdf2image.exceptions import (
    PDFInfoNotInstalledError,
    PDFPageCountError,
    PDFSyntaxError
)
import shutil
from zipfile import ZipFile
from bs4 import BeautifulSoup

#Set Google Credentials (path to JSON key) and base directory (where script is located)
os.environ["GOOGLE_APPLICATION_CREDENTIALS"]="C:/Users/Brandon/Downloads/pptx-to-vid/T2SKey.json"
base_dir = '/Users/Brandon/Downloads/pptx-to-vid'

# Create directories if they do not exist already and Define Folders
if not os.path.exists(base_dir + '/img'):
    os.mkdir(base_dir + '/img')
if not os.path.exists(base_dir + '/audio'):
    os.mkdir(base_dir + '/audio')
if not os.path.exists(base_dir + '/video'):
    os.mkdir(base_dir + '/video')
if not os.path.exists(base_dir + '/pdf'):
    os.mkdir(base_dir + '/pdf')
if not os.path.exists(base_dir + '/text'):
    os.mkdir(base_dir + '/text')
if not os.path.exists(base_dir + '/zip'):
    os.mkdir(base_dir + '/zip')
if not os.path.exists(base_dir + '/slide'):
    os.mkdir(base_dir + '/slide')

text_dir = base_dir + '/text/'
img_dir = base_dir + '/img/'
audio_dir = base_dir + '/audio/'
video_dir = base_dir + '/video/'
pptx_dir = base_dir + '/pptx/'
pdf_dir = base_dir + '/pdf/'
output_dir = base_dir + '/zip/'
extract_dir = base_dir + '/slide/'

#Function for naming images when extracted from the powerpoint.
TXT_COUNT = 1
def txt_increment():             
    global TXT_COUNT 
    TXT_COUNT = TXT_COUNT + 1

COUNT = 1
def increment():             
    global COUNT 
    COUNT = COUNT + 1

#Converts the .PPTX to a .PDF; Then Converts the .PDF to multiple renamed .PNG Files.
pptxnames = os.listdir(pptx_dir)
for i in pptxnames:
    if i[len(i)-4: len(i)].upper() == 'PPTX':

        pptx_path = pptx_dir + i
        plain_title = i[:-5]
        pdf_name = plain_title + '.pdf'
        pdf_path = pdf_dir + pdf_name
        
        if not os.path.exists(pdf_path):    #Ensures a duplicate pdf is not made.
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            powerpoint.Visible = 1

            deck = powerpoint.Presentations.Open(pptx_path)
            deck.SaveAs(pdf_path, 32) # formatType = 32 for ppt to pdf
            deck.Close()
            powerpoint.Quit()

            #Extracts images from .PNG, as well as renames them in orient "00X.png"
            images = convert_from_path('{}'.format(pdf_path), output_folder = img_dir, fmt='png', size = 1280)
            del images
                
            for f in os.listdir(img_dir):
                img_title = img_dir + f
                print(img_title)
                f_name, f_ext = os.path.splitext(img_title) 
                f_name = "Slide" + str(COUNT)
                increment()

                new_name = '{}{}{}'.format(img_dir,f_name, f_ext) 
                os.rename(img_title, new_name) 
        else:
            print('{} Already Exists!'.format(pdf_name))

        shutil.copy(pptx_path, output_dir)
        new_pptx = output_dir + i
        print(new_pptx)
        os.rename(new_pptx, new_pptx + ".zip")
        print(new_pptx)
        with ZipFile(new_pptx + '.zip','r') as zipObj:
           # Extrct all contents of zip file in current directory
           zipObj.extractall(extract_dir)
           print("extracted!")

#Create a .mp3 audio file from each .txt file.
slide_dir = extract_dir + "ppt/notesSlides/"
slidenames = os.listdir(slide_dir)
print(slidenames)
for i in slidenames:
    if i[len(i)-3: len(i)].upper() == 'XML':
        print(i)
        slide_path = slide_dir + i
        with open("{}".format(slide_path)) as fp:
            soup = BeautifulSoup(fp,"xml")
            tags = soup.find_all('a:t')
            final_text = tags[0].text
            print(final_text)
            pt_name = i[:-4] + ".txt"
            text_path = text_dir + pt_name
            text_file = open(text_path , "w")
            n = text_file.write(final_text)
            text_file.close()

for f in os.listdir(text_dir):
        text_title = text_dir + f
        print(text_title)
        f_name, f_ext = os.path.splitext(text_title) 
        f_name = "Slide" + str(TXT_COUNT)
        txt_increment()

        new_name = '{}{}{}'.format(text_dir,f_name, f_ext) 
        os.rename(text_title, new_name)
           
textnames = os.listdir(text_dir)
for i in textnames:
    if i[len(i)-3: len(i)].upper() == 'TXT':
        
        text_path = text_dir + i
        plain_name = i[:-4]
        audio_name = plain_name + '.mp3'
        audio_path = audio_dir + audio_name
        
        if not os.path.exists(audio_path):
            #Google Text-to-Speech API Processing
            raw_text = open('{}'.format(text_path),"r").read()
            client = texttospeech.TextToSpeechClient()
            synthesis_input = texttospeech.types.SynthesisInput(text=raw_text)

            voice = texttospeech.types.VoiceSelectionParams(
                language_code='en-US',
                ssml_gender=texttospeech.enums.SsmlVoiceGender.NEUTRAL)
            
            audio_config = texttospeech.types.AudioConfig(
                audio_encoding=texttospeech.enums.AudioEncoding.MP3)
            response = client.synthesize_speech(synthesis_input, voice, audio_config)

            #Places Audio File into 'Audio' Folder
            with open('{}'.format(audio_path), 'wb') as out:
                out.write(response.audio_content)
                print('{} Created!'.format(audio_name))
        else:
            print('{} Already Exists!'.format(audio_name))

#Combine .PNG and .MP3 to create Individual .MP4s
imgnames = os.listdir(img_dir)
for i in imgnames:
    img_name = i
    plain_name = img_name[:-4]
    img_path = img_dir + img_name
    audio_name = plain_name + '.mp3'
    audio_path = audio_dir + audio_name
    
    video_name = plain_name + '.mp4'
    video_path = video_dir + video_name
    if not os.path.exists(video_path):
        command = 'ffmpeg -loop 1 -i {} -i {} -c:v libx264 -tune stillimage -c:a aac -b:a 192k -pix_fmt yuv420p -shortest {}'.format(img_path,audio_path,video_path)
        subprocess.run(command, shell = True)
        print("{} Complete!".format(video_name))
    else:
        print("{} Already Exists!".format(video_name))

#Create a list text file of all videos
final_video = base_dir + '/output.mp4'

if not os.path.exists(final_video):
    videonames = os.listdir(video_dir)
    vid_list = open("vid_list.txt", "w+")
    vl_path = base_dir + '/vid_list.txt'

    for i in videonames:
        video_path = video_dir + i
        vid_list.write("file '{}'\n".format(video_path))
        
    vid_list.close()

    #Use FFmpeg to concatenate all the videos together into a single .MP4 File
    command3 = 'ffmpeg -f concat -safe 0 -i {} -c copy output.mp4'.format(vl_path)
    subprocess.run(command3, shell = True)
    os.remove('{}'.format(vl_path))
    print("Video is Complete! Placed in output.mp4")
else:
    print("Final Video has already been made!")
