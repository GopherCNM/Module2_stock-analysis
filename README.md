# VBA of Wall Street

## Overview of Project

### Purpose

The purpose of this analysis is to help our client, Steve, to create a scalable stock analysis report using VBA. He wants to be able to run the report for several different stocks for a specified year, and then pull back information regarding each stock’s volume and return over that year. Prior to this, we created an analysis that did this analysis for us, but now we’re looking to improve (or refactor) our code to enhance its performance. Steve will use this reporting to help his parents make investing decisions.

## Analysis and Challenges

### Results & Analysis: Stock Performance

Regarding stock performance for the years reviewed, 2017 showed better returns for this portfolio of stocks than did 2018. Only 2 stocks, ENPH and RUN, showed positive returns for both years. Steve’s parents also wanted to look at stocks with material trade volume, and both stocks satisfy this as well. For the 12 tickers reviewed, ENPH had 5th most volume in 2017 and 1st in 2018, and RUN had 4th most volume in 2017 and 3rd most in 2018. Based on this analysis, both of these stocks may be worth further consideration.

### Results & Analysis: Code Performance

In the refactored code for this challenge, we created output arrays to capture volume, starting prices, and ending prices for each of the tickers reviewed. This allowed us to close that loop before moving on to part 4 of the challenge, where we dropped the values into our output sheet. Qualitatively, it felt more efficient than creating the nested loops that we did in the asynchronous material. Our timer test also supported this quantitatively. The results of the timer show that our refactoring shaved roughly 0.5 seconds off the code’s run time.

***2017 Run Time***
![2017 Refactored Run Time](/Resources/VBA_Challenge_2017.png)
![2017 Old Run Time](/Resources/VBA_Challenge_2017_old.png)

***2018 Run Time***
![2018 Refactored Run Time](/Resources/VBA_Challenge_2018.png)
![2018 Old Run Time](/Resources/VBA_Challenge_2018_old.png)

### Challenges and Difficulties Encountered

This was my first experience with VBA. For the most part, the challenge went smoothly. After my first attempt, I was receiving an “overflow” error message when I tried to run the code. The debug was pointing me to a part of the code that wasn’t ultimately the source of the issue. I used Google to research the issue and possible causes, and also reached out to my classmates on Slack. I realized that I have over-complicated the refactored code by creating a nested loop and inadvertently using a variable name twice.

Working through this taught me the importance of a few things:
-the value of keeping things simple when writing and refactoring code
-don’t hesitate to rely on resources such as Google and the experience of others
-the collaborative nature of coding
-not always relying on Excel’s debugger to point you to the source of the coding issue

## Summary

- What are the advantages and disadvantages of refactoring code?

1. **Advantage:** Refactoring allows you to critically review existing code with another set of eyes. After using existing code and seeing how it performs, refactoring allows you to give it another pass and think about ways to structure things that make more sense.
2. **Advantage:** Refactoring may help to enhance performance by simplifying code or enhancing its efficiency. This may translate to speed to speed and less strain on system resources. In our exercise, these speed differences were small, but over large data sets I would imagine refactoring could help to save considerable time.
3. **Disadvantage:** I could see refactoring being risky if the person doing the refactoring doesn’t fully understand the intent or structure of the existing code. Why was it set up this way and what is it trying to do? How does it currently perform? Without this, refactoring could do the opposite of what’s intended, or break the existing code.

- How do these pros and cons apply to refactoring the original VBA script?

Most of the above advantages and disadvantages came into play in this refactoring exercise for this challenge. I temporarily broke functioning code, but then came to a better understanding of how it worked. Ultimately, by refactoring the existing code, I was armed with more information and knowledge than I had when I created the initial code. By refactoring it, I improved the speed and also made more detailed notes within the code to help myself or anyone who may later review it.
