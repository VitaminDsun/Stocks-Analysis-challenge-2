# Stocks-Analysis-challenge-2
stock analysis challenge 2

#Overview 
I was brought in to help Steve make a few analyzations for his parents on some stocks in the years of 2017 and 2018. they are interested in investing in green energy and believe it will be a good investment for the future, they originally were interested in a company named DAQO New Energy Corp (ticker ID DQ). DQ makes silicon wafers for soler panels, so it fits their want to invest in to green energy, and they also meet at dairy queen (what a cute story). although after some investigation it was found that they might not be the best investment after we ran some number on DQ alone, this led Steve to the idea of diversifying their portfolio with other green energy companies. To do this we took the code that we used to find the information on DQ can be found on the (DQ Analysis All Stocks Analysis) and refactored it to be used to find the ticker volume, the starting price and the ending prices for some other stocks provided, we then took the opportunity to refactor the code so that it could be used on much larger data sets in other analysis this was done by taking the code I created for the DQ analysis (Modules 1, 5 in the VBA Macros ) and was able to consolidate it down to a more easy to read and easier to use code (Module 2 in the VBA Macros.)

results before refactoring 
2017
![Screen Shot 2022-01-16 at 9 21 09 PM](https://user-images.githubusercontent.com/93777016/149706667-3b51d3dd-5537-4df1-bfb3-fd682c0e94b0.png)
2018
![Screen Shot 2022-01-16 at 9 19 22 PM](https://user-images.githubusercontent.com/93777016/149706683-73ac1fd6-5ec7-48cb-8c01-b80d56351a34.png)

results after refactoring 
2017
![Screen Shot 2022-01-16 at 9 30 30 PM](https://user-images.githubusercontent.com/93777016/149706728-26db46d5-9865-4730-b09c-28b5b43ab56e.png)
2018
![Screen Shot 2022-01-16 at 9 30 14 PM](https://user-images.githubusercontent.com/93777016/149706732-99663d5b-1b12-4664-b514-dde251e487e2.png)


#Summary
 Overall Steve’s happy, so that makes me happy!
What are the main advantages or dis-advantages of refactoring code?
Advantages-
Refactoring code is changing the source code without modifying the existing code’s main function, (saves a lot of time when refactoring code is done correctly) it has many advantages such as improving the codes readability and by reducing complexity of the code. This allows for the code to be easily maintained and built upon in the future by other developers. Another way refactoring code can be an advantage is if done correctly it can improve a codes functionality by reducing duplications within code and simplify the code to make it easier to understand. This will help make the logic of the code standout another advantage of proper refactoring code will be clear and defined layers of code with proper uses of whitespace and labeling of what the codes responsibility and functions; As well as any changes that were made in comments. This is also a good learning opportunity because everyone can look at a problem in a different way and this will be a good opportunity to see how someone else might have solved the same issue you are asking yourself now. 
	
Dis-advantages-
The main dis-advantages with code refactoring is working with the previous code itself, this is because when you start with a code that’s already written and if it was not maintained well or you don’t understand the logic the code is trying to use. It can sometimes be a mess and more of a hassle to deal with then just starting from scratch (Wasting a lot of time), But my grandfather used to say if it’s not broke don’t fix it. So sometimes just finding the pieces of the code that does work can some time point you in the right direction to the reason why it might be having issues or could be improved upon. Another disadvantage is you might break the code if you don’t know what you yourself are trying to do with the code; or what the code is trying to do. This can happen by trying to remove what you might think is dead code or unused and in fact it might be holding a major piece together.

@How do these pros and cons apply to refactoring the original VBA script?
Refactoring the original code on VBA-
When refactoring a code into VBA some of the pros and cons I listed above happened during my time working with this code, by refactoring the code with the starter code I was able to remove un-needed code and was able to organize my code better. But a lot of confusion happened during the refactoring code portion of this project. This has also helped me identify some of the things in my own code that might not have made as much sense to someone as it did me in the moment. Once I was able to get a few examples of how to refactor the code a bit cleaner I was able to see where I was repeating myself a lot, as well as not doing the whole code on one module this was very confusing for myself to work with; I thought it would make things easier to have them on different windows but in the end, I was able to condense it down to one window. This allowed me to be able to look at the whole code without having to switch back and forth to see what I was doing on each of the macros. This is one way that refactoring my original code on VBA had some pros and cons, after the refactoring of the code the code itself ran so much faster and having the whole code on one sheet allowed the gathering of the data and the formatting to happen all at once. before when I had it on two sheet you would have to analyze the data, and then you would have to format it. But after the refracting the code everything happened at once, which in theory saved me even more time.



