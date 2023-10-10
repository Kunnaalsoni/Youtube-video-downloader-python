import pandas as pd
import numpy as np
from operator import index
import urllib.request
import pandas as pd
from pytube import YouTube
import os
try:
    read_file = pd.read_excel('links.xlsx')
    print("File read successfully")

except Exception as e:
    print("File parsing error")
    print(e)
    exit()

for x in read_file.index:
    try:  
        # object creation using YouTube 
        # which was imported in the beginning  
        yt = YouTube(read_file['Link'][x])
        youtubeObject = yt.streams.filter(progressive=True,file_extension='mp4').order_by('resolution').desc().first()
    except Exception as e:  
        print("Connection Error") #to handle exception  
        print(e)

    try:
        #downloading the video  
        d_video = youtubeObject.download()
    except Exception as e:
        print("An error has occurred")
        print(e)

    print('Task Completed!')  


# y=1
# for x in read_file.index:
#     yt = YouTube(read_file['Link'][x])
#     yt.title
