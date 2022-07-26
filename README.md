# Barcode-Scanning-System #


## Problem Statement: ##

Aim of the project is to scan the barcodes (1D &amp; 2D) and store the data associated to the barcode in the excel sheet. This project is developed in python. 
The reason for writing this program in Python is that the necessary commands for Barcode Scanning are accessible in Python, and the Python programming language is simple. Python has a strong & powerful tool that has different toolboxes which are of great help in Digital Image Processing.

## Libraries Used: ##

1. OpenCV-Python:<br />
   - This python library binds designed to solve computer vision problems.
   - cv2.VideoCapture() lets you create a video capture object which is helpful to capture videos through webcam and then you may perform desired operations on that     video.

2. Pyzbar:<br />
    - The main feature of pyzbar is that decoding is done directly using the same function. In addition, it supports decoding multiple barcodes or QR Codes in a single image.

    - decode() function returns a list of namedtuple called Decoded. Each of them contains the following fields:

      - data — The decoded string in bytes. You need to decode it using utf8 to get a string.

      - type — Only useful for barcodes as it outlines the barcode format.

      - rect — A Rect object which represents the captured localization area.

      - polygon — A list of Point instances which represents the barcode or QR Code.

3. Openpyxl:<br />
    - The openpyxl module allows Python program to read and modify Excel files. 
    - With help of this i created an excel file in the system and created table in same file listing data and type of each barcode that is scanned. 

## Results/Interface: ##

<p align="center">
  <img src="https://user-images.githubusercontent.com/56083892/180940024-d66a0dd4-a0c3-4757-8040-9338b8e3dbae.png">
</p>

<h2 align="center">Flowchart of the Program</h2><br/><br/>


 ![1](https://user-images.githubusercontent.com/56083892/180941458-442be529-a882-4a58-b67b-4974e842ea40.png)

 <h2 align="center">Reading Code128 1D Barcode</h2><br/><br/>

 

 ![2](https://user-images.githubusercontent.com/56083892/180941641-aad7e50e-d93c-4143-9d61-b591737d36e1.png)

<h2 align="center">Reading EAN13 1D Barcode</h2><br/><br/>


![3](https://user-images.githubusercontent.com/56083892/180941656-f42edf81-e009-4296-9d24-347279121bb7.png)

<h2 align="center">Reading Code128 1D Barcode</h2><br/><br/>



![4](https://user-images.githubusercontent.com/56083892/180941667-54a89cd3-a470-45b7-a058-96f3a8875a40.png)

<h2 align="center">Reading QR 2D Barcode</h2><br/><br/>



<p align="center">

  <img src="https://user-images.githubusercontent.com/56083892/180942265-eaec3cb0-d473-4952-ab7a-947dd2d1fcf8.png">

</p>

<h2 align="center">Barcode Data stored in Microsoft Excel</h2><br/><br/>
