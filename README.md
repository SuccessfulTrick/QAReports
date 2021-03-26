# QAReports - Manufacturing Quality assurance Report

## where did the idea of the project come from ? GenerateDIMS.py

We live in a very fast moving world with changing technology even tho we still have very experienced people,
that can do amazing work in different roles that doesn't require Computers or they lack technical knowledge,
or does not have the ability to use them,we always need to find new and creative ways to keep up and find solutions,
After the step of auditors examinations role, Comes the technical part of generating a report which is done by entering the data manually,
and that needs basic level of technical knowledge which not every person in these roles has,
but also can be draining and lead to many Mistakes, Copy and Pasting, such as the same special symbol for different
dimensions especially when there are hundreds of dimensions in 1 Drawing.

And from there comes this project and other more ideas which is still in early stages and under development.

## Main Goal of QAReports Project

The main goal of this project to keep developing new ways and ideas to Save mistakes,time,
and most important is to make it as easy as possible to generate a final report with the right template
for the right customer by any person that lacks technical knowledge, working knowledge, or any other barriers.

Ideas as:  
• Scanning and extracting the dimensions from the scanned pdf  
• Extracting the relevant drawing name, ID and information from the ERP System  
• Automatically filling required dimension, tolerances and results in the required QA Report  
• Detecting which QA Final report template is relevant for the specific drawing and customer  

Example of a finished word document report relevant to the drawing given as an example.
( Balloons 4-6 were generated just as an example use of the program they're not in the drawing )

![alt text](https://i.gyazo.com/63c72dfcf4050947c3d881d4f21f6096.png)

## What's the proccess before generating report ?
*Images are part of the full drawings, are relevant to this project only.

Steps From Receiving technical drawing to auditors examinations: 

![alt text](https://i.gyazo.com/aee9080d90d8d3dcdd1e4e19e15941b2.png)

Example of a technical drawing: 

![alt text](https://i.gyazo.com/8f12cdf5befc1bd3383bdc255b52db12.png)

Drawing after examination with results:

![alt text](https://i.gyazo.com/d5362dba1aba0b3cadca8893f50362b0.png)

## So what does the program do?

To this moment the program returns an Excel table with the required data that is needed to fill in the QA Final report template,
includes checking the sample size needed for the serving quantity given, with a random dimension for each item
that was generated from the Minimum and maximum dimensions for the required dimension given.

Data inputs Entered in Pycharm :
![alt text](https://i.gyazo.com/c202a94497cf2933789bf57fc587d71a.png)

Excel Table generated : 

![alt text](https://i.gyazo.com/70cb77588f43d57d003d15d7d4c7bf9b.png)
