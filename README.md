# VBA MODULE

### On slack I saw some things shared, especially the importance of the third point I put on here. The credit card code from last thursday was crucial.

1. A very important thing I learned was this as a first source:
    https://stackoverflow.com/questions/39581487/loop-through-all-worksheets-in-workbook/53673884#53673884

2. I learned I could dim as Worksheets and iterate through them all. 
    I only had to use this correctly and make sure my values were set at the correct time. 
    It became a problem when I forgot to set values when doing this worksheet iteration.

3. Another HUGE source was the credit card iteration code
    I changed it to fit this scenario and loop to give me the first table.
    This boils down to rewatching the Zoom classes, that's where I got the code for the 'lastrow' on my file by using xlup

4. I wasn't giving me the results I needed so I saw Yash in tutoring as another source
    I had the right idea but was just beginning to play with it.
    I have to say I realized how little variables I had during this tutoring lesson.

5. Things that were unknown to me and I had to look up, were the following:
    -Formatting: cells as percentages and colors (line79, 82-87) the conditionals were easy but I had to look up the formatting.
    -Searching for the greatest % value. 
        -I didn't know how it relates to a regular integer so I had to use AI to ask about it and there's no difference it seems
    
6. I kept on making a mistake on the small table: 
    -It kept on running the lowest and biggest 'Percentage Change' of all the worksheets totaled until it occurred to me to set a value BEFORE the worksheet iteration started 
            # THIS WAS A HEADACHE


