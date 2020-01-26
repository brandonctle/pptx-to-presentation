## Inspiration
As a student technician for the Texas A&M Engineering Department of Remote Education, the most pressing issue of future education I learn is that Professors are having to increase their content to students with less and less time, especially with the increased digitation of education. 

## What it does
In this project, we obtain a client-side PowerPoint (**.pptx**) file with a transcript of notes located in the annotations section.

From the PowerPoint, we use Python to extract the annotations (as text files) and the slides (as images), translate the annotations from a text file to an audio file using the Google Cloud Machine Learning's Text-to-Speech API, and combine the audio files and images to create a single video file (MP4) of the presentation.

## How I built it
## Installation
This Python Program requires multiple pre-installed software to run:
*Note: if you haven't installed [Python], it is the base coding language of the project*

#### Google Cloud Project and SDK
Setup for Google SDK can be found [here].
The SDK is what allows the access token for the Google Service Account to be read and successfully processed. Only through installing and setting up the SDK will the Text-to-Speech API work.

Please be sure a project is created in **Google Cloud Console**. 
- The Google Text-to-Speech API must be enabled.
- A service account must be created with the API
- A client key (JSON) must be downloaded and placed in a path file that can be found.

Additionally, you'll need to `pip` install the Google API:

    pip install --upgrade google-cloud-texttospeech

#### Addtional Dependencies
This code relies on multiple different requirements to run.

- [FFmpeg](https://www.ffmpeg.org/) - Responsible for combining audio and images, as well as concatenating (combines) videos together.
- [Pdf2Image](https://github.com/Belval/pdf2image) - Extracts X amount of images for a PDF files with X pages.
    - [Poppler](https://poppler.freedesktop.org/) *Required for P2I*
- [comtypes](https://pypi.org/project/comtypes/) - *Windows Only* - Responsible for PPTX-to-PDF Extraction.

*The links above provide installation and documentation. Additionally, set both FFmpeg and Poppler's `/bin` folder into the `PATH` Variable (as described in each application's documentation.*

---
## Setup
In order to properly run the script, the directories of the Credentials Key and the location of the project folder must be setup.

#### Credentials Path 
The `GOOGLE_APPLICATION_CREDENTIALS` must be set to the path of the JSON file. To set it while running a Python Script, the setting must be set in the python code.

```sh
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "[INSERT JSON PATH HERE]"
```

#### Setting Directories
A base directory also needs to be set. In the project folder, there should be a subset folder titled "pptx", which contains the PowerPoint to be converted.

```sh
    base_dir = '[PATH TO PROJECT FOLDER]'
```
*Note: This pathway should go to the folder that *contains* the PowerPoint folder, not the actual Powerpoint itself.*

## Extracting Annotations from PowerPoint
In order to extract text from the annotations of the PowerPoint, the presentation must be converted to its raw files (**XML**) and manually pulled out of the files. In order to accomplish this, we utilize two libraries: **Zipfile** and **BeautifulSoup**

*Zipfile* converts the PowerPoint to a zip file, where the raw data files can then be extracted.
*BeautifulSoup* extracts the text out of the annotations of the PowerPoint files.

## Creating the Video File
### Initialization
The python script begins by checking for, creating, and defining 5 subdirectories: 
```
/pptx/ -- Should already be created and contain the MP4 Files.
/pdf/
/img/
/text/
/audio/
/video/
```
These subdirectories are automatically created in the designated location established by the `base_dir` input.
### The Increment Function
A single Incrementing Function Helps to Name our Images once they have been extracted from the powerpoint. It is later referenced.
### Converting PowerPoint to Images
There is (currently) no automatable way to easily export PNG files from PowerPoint in a one-step process in Python.
In order to create an Image from each slide of a PowerPoint Presentation, we must use convert it to a pdf, where Python then has a library which can convert a PDF into PNG Images.
##### The Comtypes Client
*"comtypes allows to define, call, and implement custom and dispatch-based COM interfaces in pure Python"*
For this project, comtypes is used to open PowerPoint at the specified location of the file, save the file as a **pdf** format, and close PowerPoint.
##### Using pdf2image
Now that we have a pdf, we can use the pdf2image Python Library to extract and save each 'page' of the pdf - which is a 'slide' of PowerPoint - as an image.
We do so using the `convert_from_path` keyword. Additionally, we need to specify parameters provided in the `convert_from_path`, including `output_folder`, `format`, and `size`.

The images extracted also have a unique identifier name which needs to be renamed in order to be merged with the text/audio. In this case, we rename the files to **OOX.png** by file order. It doesn't matter what the files are named, so long as they are in order, and match the names of the text files.

***Important Note: The PNG's extracted MUST have dimensions thta are divisible by 2 (even). They will not be able to be converted to video otherwise - which is why we use size as a parameter to ensure corect formatting.***

### Converting Text to Audio
The primary tool used in the conversion process is **Google Cloud Text-to-Speech API**, a machine learning tool. Various parameters and configurations can be adjusted, including the language, gender, output type, and even how fluent the text sounds (there is a premium, more expensive option for the translation process to have a more "natural" sounding voice). For this project, settings are kept as default and exported as a MP3 file.

##### Costs of Translation

### Combining Image and Audio to Video
At this point, we have a set of Images and Audio files. The total amount should be equal to each other, and each image should correspond to their correct audio file.

FFmpeg, with a single line command, creates an MP4 Video File.
`ffmpeg -loop 1 -i {image} -i {audio} -c:v libx264 -tune stillimage -c:a aac -b:a 192k -pix_fmt yuv420p -shortest {video}`
By the end of the process, you should have a set of MP4 files which match the amount of slides the powerpoint contains.

### Creating the Complete Video
Finally, another FFmpeg command is used to concatenate all the videos into a single MP4 file.
`ffmpeg -f concat -safe 0 -i {list file} -c copy output.mp4`
This is performed by automatically creating a "list file" with all the names of the separate MP4 files, finding those files, and merging them together.

The Final Video is Outputted as **"output.mp4"**, located in the same directory as the Python Script.

## Challenges I ran into
This program initially used various Python libraries, including comtypes, a **windows only** python library. Therefore, the complete program will NOT work on any other operating system.

Searching through libraries and code which would be cross-platform compatible was difficult, but ultimately possible. Where libraries didn't support the code, it was subsidized with code.

Google Cloud's Text-to-Speech API is also one constantly developing, so figuring out how to implement a server style code into a clientside program was difficult, but eventually doable after reading documentation as well as trial-and-error processes.

## Accomplishments that I'm proud of
I'm extremely proud of Google Cloud integration and Machine Learning use. Additionally, I'm proud to have the entire program run on its own, without the need for user interaction the entire time.

## What I learned
I learned about Python library integrations, Google Cloud Machine Learning client uses. Additionally, I was able to learn more about operating system modifications.

## What's next for PowerPoint to Presentation

The next hope is to improve the flexibility of the language and utilization of the program. These improvements include:
- Improving the Machine Learning API to understand the text more flexibly.
- Streamlining the extraction process to be faster and more efficient.
