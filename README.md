# All Stock Analysis Utilizing VBA

## Overview of Project

Project initiated to analyze multiple stock trends in 2017 and 2018 utilizing VBA. 

## Purpose

Provide client with a streamlined way to evaluate trends in stocks. Focusing on total daily volumes and percentages of return in particular years. 

## Analysis and Results

In order to complete the analysis and ensure the performance of the code was running smoothly the original code developed was refactored. By doing so the refactored code was made to loop through all of the data one time and gather all requested information needed for the analysis. 

### Analysis

Original code analyzed each stock independently using a nested for loop which is shown below. By refactoring the code utilizing a ticker index variable that then collected data for each variable we were able to improve the overall efficiency while maintaining the integrity of the data.

#### Original Code vs Refactored Code

##### Original 

<img width="400" alt="Original_Code" src="https://user-images.githubusercontent.com/102195085/167179936-5c6c2050-999a-49c6-b1e6-aad7ca5b1ecf.png">    

##### Refactored

<img width="400" alt="Refactored_Code" src="https://user-images.githubusercontent.com/102195085/167179954-610a5831-6fb8-4a3d-af7b-26f1d381c2ca.png">

### Results

By refactoring the original code the efficiency improved across the board. For example in the analysis of the 2017 stocks the code was able to run in 1/3 the time after being refactored. Results of the two years are shown below.

#### 2017

##### Original

<img width="400" alt="VBA_Challenge_2017_Original" src="https://user-images.githubusercontent.com/102195085/167182658-9263bea1-e9da-4fab-be67-c32f15b49299.png">    
##### Refactored 

<img width="400" alt="VBA_Challenge_2017_Refactored" src="https://user-images.githubusercontent.com/102195085/167182695-d48df0c2-4ec5-4f57-9c13-5bbaf6f95f92.png">

#### 2018

Original 
<img width="400" alt="VBA_Challenge_2018_Original" src="https://user-images.githubusercontent.com/102195085/167182777-6cd90565-88b0-4439-9e88-54f2aa37b871.png">    
Refactored 
<img width="400" alt="VBA_Challenge_2018_Refactored" src="https://user-images.githubusercontent.com/102195085/167182785-2e8788bb-cd15-42b2-93f3-3c5ee2de3fe0.png">

## Summary

Throughout this project the goal was to deliver a program that would allow the client too quickly and effectively evaluate certain stock performance. This was achieved by writing code utilizing VBA and outputting it in a clear and easy to follow excel sheet. By refactoring code you can make code run more efficiently and utilize less resources. Refactoring improved the overall efficiency of the program by cutting the run time by 2/3 of its original. While this is a great advantage some of the disadvantages are that the code was more difficult to write and appears to be more cumbersome than the original code. With that said as it can handle more data in the same or less time the advantages of refactoring far outweigh the disadvantages.
