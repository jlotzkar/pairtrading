# -*- coding: utf-8 -*-

"""

Created on Fri Jun 18 00:53:53 2021

 

@author: Lotzkar

"""

 

import streamlit as st

from PIL import Image

#import cv2

import base64

 

st.set_page_config(

    page_title = "CIBC Streamlit",

    layout = "wide")

 

#containers

header = st.beta_container()

dataset = st.beta_container()

features = st.beta_container()

 

with header:

    st.title("Pair Trading Analysis Dashboard")

    st.write("Below are the two picked equities and a series of graphs showing their price histories and buy/sell points on the assumption of mean reversion.")

 

#from meanreversionedited import *

 

with dataset:

    st.header("Plot of historical stock prices for the two selected securities:")

   

#image = Image.open('C:/Users/lotzkar/Desktop/Mean Reversion Project/plot1.png')

#st.image(image, width = 800, output_format='png')

 

from pathlib import Path

 

def img_to_bytes(img_path):

    img_bytes = Path(img_path).read_bytes()

    encoded = base64.b64encode(img_bytes).decode()

    return encoded

 

header_html = "<img src = 'data:image/png;base64,{}' class='img-fluid'>".format(

    img_to_bytes("plot1.png")

)

 

with features:

    st.markdown(

    header_html, unsafe_allow_html = True,

)

 

header_html_2 = "<img src = 'data:image/png;base64,{}' class='img-fluid'>".format(

    img_to_bytes("plot2.png")

)

   

with features:

    st.header("Plot of rolling z-score for price ratio between two selected securities:")

    st.markdown(

    header_html_2, unsafe_allow_html = True,

)

   

header_html_3 = "<img src = 'data:image/png;base64,{}' class='img-fluid'>".format(

    img_to_bytes("plot3.png")

)

   

with features:

    st.header("Plot of historical stock prices for the two selected securities with buy/sell signals:")

    st.markdown(

    header_html_3, unsafe_allow_html = True,

)

 

header_html_4 = "<img src = 'data:image/png;base64,{}' class='img-fluid'>".format(

    img_to_bytes("plot4.png")

)

 

with features:

    st.header("Today's signal:")

    st.markdown(

    header_html_4, unsafe_allow_html = True,

)  