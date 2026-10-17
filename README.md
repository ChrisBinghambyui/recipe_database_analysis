# Overview

As a regular internet user, I have seen many sites with useful filtering tools and navigation helpers. However, I have yet to find one that scratches that itch with recipes. I've wanted to make one for a while, and now, I have! Welcome to PantryPal, a little tool to help home cooks organize their pantries and find ideas for new meals.

Thanks to a massive recipe database from Kaggle, I was able to get a baseline started, but I'm planning on allowing users to enter their own recipes in the future. Users can search and filter based on ingredients, seeing what they can make with what's readily available, or what they need to buy to fulfill requirements for a recipe. The 'Best Buys' tab lists out popular ingredients that will open up the option to make many more recipes.

If you have any feedback, I'd really appreciate it if you took the time to fill out my feedback survey [here.](https://docs.google.com/forms/d/e/1FAIpQLSfwoFzmMW6um5t2pzUnB7f0iB8C0iP2R_4OIiHqTUvzZG5W6A/viewform?usp=header)

{Provide a link to your YouTube demonstration.  It should be a 4-5 minute demo of the data set, the questions and answers, the code running and a walkthrough of the code.}

[Software Demo Video](http://youtube.link.goes.here)

# Data Analysis Results

**Question 1: With a typical pantry, how many recipes can I make right now, and how much does the matching rule change that number?**

Each recipe is compared against the pantry one ingredient name at a time. Under strict matching, an ingredient counts as covered only when every word of its name appears in a single pantry item. Under loose matching, a pantry item also covers any longer ingredient that contains it, so "chicken" covers "chicken breast". I counted the recipes with nothing missing under both rules using my own pantry list. Loose matching finds more recipes, but it also lets through some bad matches, such as butter standing in for peanut butter. Strict matching is the safer default, and the loose option is there for people who want a longer list and are ok with a little manual work.

**Question 2: Which single ingredient would unlock the most new recipes if I bought it?**

This is the Best Buys tab. The program collects every recipe that is missing exactly one non-seasoning ingredient, then counts how often each missing ingredient appears. The ingredient with the highest count is the best purchase. For my pantry, the top result was [ingredient], which would complete [N] recipes. The ranking showed me that a handful of common items account for a large share of the near misses.

Techniques used:

* **Filter:** title, must-contain, and cannot-contain filters, plus the pantry check itself.
* **Sort:** recipes are ordered by how many ingredients are missing, and every column in the results tables can be sorted by clicking its heading.
* **Aggregate:** counts of makeable recipes per tab and a frequency count of missing ingredients.
* **Data conversion:** the dataset stores lists as JSON text, so each row is parsed, lowercased, stripped of filler words such as "chopped", and reduced to singular forms so that "tomatoes" and "tomato" match.

# Development Environment

{Describe the tools that you used to develop the software}

{Describe the programming language that you used and any libraries.}

# Useful Websites

{Make a list of websites that you found helpful in this project}
* [Web Site Name](http://url.link.goes.here)
* [Web Site Name](http://url.link.goes.here)

# Future Work

{Make a list of things that you need to fix, improve, and add in the future.}
* Item 1
* Item 2
* Item 3