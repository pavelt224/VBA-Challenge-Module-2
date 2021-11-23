# VBA-Challenge-Module-2
Assignment for Module 2

# Purpose
In this project and analysis, we'll use VBA solution code to alter or reorganize the Stock Market Dataset so that we can collect all of the data at once. Then we'll see if restructuring the code resulted in the VBA script running faster. Finally, we want to make the code more efficient by reducing the number of steps, utilizing less memory, or enhancing the logic to make it easier to read for future users.

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

## How do these pros and cons apply to refactoring the original VBA script?

Code refactoring is the process of improving or updating code without affecting the software's functionality or external behavior. Now consider this: what happens if you need to troubleshoot 
your code after a few days or months? Is it difficult? Is it difficult to comprehend? If you answered yes, you most likely did not pay attention to improving or reorganizing your code.

