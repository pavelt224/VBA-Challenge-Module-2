# VBA-Challenge-Module-2
Assignment for Module 2
Advantages:

In well-structured code with nested conditionals and loops, logical mistakes are common.
In our situation, Excel flow makes program logic more understandable by separating it from the sequence in which the underlying code is created.
VBA code interpretation (Excel) can uncover patterns that aren't visible in the raw code.

Disadvantages:

You can adjust the logic to delete duplicate lines in a long process that has the identical line of code in numerous places.
It's possible that a logical structure is replicated in two or more procedures (possibly via copy & paste coding). When this logic is discovered, it is better to relocate it to a new function and call it from the other functions.
It's usually advisable to separate a complex unstructured code into numerous functions.
The testing results may be influenced by the refactoring process.

How do these pros and cons apply to refactoring the original VBA script?

Code refactoring is the process of improving or updating code without affecting the software's functionality or external behavior. Now consider this: what happens if you need to troubleshoot 
your code after a few days or months? Is it difficult? Is it difficult to comprehend? If you answered yes, you most likely did not pay attention to improving or reorganizing your code.

