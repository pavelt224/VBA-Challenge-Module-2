# VBA-Challenge-Module-2
Assignment for Module 2

# Purpose
In this project and analysis, we'll use VBA solution code to alter or reorganize the Stock Market Dataset so that we can collect all of the data at once. Then we'll see if restructuring the code resulted in the VBA script running faster. Finally, we want to make the code more efficient by reducing the number of steps, utilizing less memory, or enhancing the logic to make it easier to read for future users.



# Results: 
Running our comprehensive 2017 and 2018 data stock analysis gave us an elapsed run time for each year. For each "stock of interest," the original (asynchronous) code iterated over the entire year's worth of stock data once. It took twelve consecutive iterations over the whole year's data to accomplish this. The following are the results for each of the two years studied::

<img width="209" alt="Initial Analysis 2017" src="https://user-images.githubusercontent.com/93852380/143138145-ac5eda85-41b4-4dc5-8bc6-a9fafd644c4a.png">

<img width="234" alt="Refactored Analysis 2017" src="https://user-images.githubusercontent.com/93852380/143138173-b6a59288-ff00-45bf-b137-c12b301a51fa.png">

The percent decrease in runtime was calculated as follows (using 2017 data)
Overall decrease = original time (asynchronous time) - new time (synchronous time) 
0.75 - 0.1308594 = 0.6191406

<img width="233" alt="Initial Analysis 2018" src="https://user-images.githubusercontent.com/93852380/143138190-e12f5915-cece-4ce1-af29-f0b4e0783154.png">

<img width="233" alt="Refactored Analysis 2018" src="https://user-images.githubusercontent.com/93852380/143138249-e2c69d8e-935d-405b-8b71-4ceebc4e8704.png">

The percent decrease in runtime was calculated as follows (using 2018 data)
Overall decrease = original time (asynchronous time) - new time (synchronous time) 
0.7578125 - 0.1298828 = 0.6279297

# Initial Analysis
<img width="341" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/93852380/143135186-d036d898-a5cd-4f42-ad50-40dc41578fde.png">


<img width="346" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/93852380/143135225-e90bc869-ddfa-44d0-8aba-adce2b793ac7.png">



# Analysis Refactored: 

<img width="357" alt="VBA_Challenge_2017(Refactored)" src="https://user-images.githubusercontent.com/93852380/143135595-8925e84c-9450-440d-8681-050d7fd626c4.png">

<img width="352" alt="VBA_Challenge_2018(Refactored)" src="https://user-images.githubusercontent.com/93852380/143135634-7b65fdd7-4e83-453d-880f-390e0e787b33.png">



# Advantages:

In well-structured code with nested conditionals and loops, logical mistakes are common.
In our situation, Excel flow makes program logic more understandable by separating it from the sequence in which the underlying code is created.
VBA code interpretation (Excel) can uncover patterns that aren't visible in the raw code.

# Disadvantages:

You can adjust the logic to delete duplicate lines in a long process that has the identical line of code in numerous places.
It's possible that a logical structure is replicated in two or more procedures (possibly via copy & paste coding). When this logic is discovered, it is better to relocate it to a new function and call it from the other functions.
It's usually advisable to separate a complex unstructured code into numerous functions.
The testing results may be influenced by the refactoring process.

