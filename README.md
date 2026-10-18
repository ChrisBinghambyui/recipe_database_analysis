# Overview

As a regular internet user, I have seen many sites with useful filtering tools and navigation helpers. However, I have yet to find one that scratches that itch with recipes. I've wanted to make one for a while, and now, I have! Welcome to PantryPal, a little tool to help home cooks organize their pantries and find ideas for new meals.

Thanks to a massive recipe database from Kaggle, I was able to get a baseline started, but I'm planning on allowing users to enter their own recipes in the future. Users can search and filter based on ingredients, seeing what they can make with what's readily available, or what they need to buy to fulfill requirements for a recipe. The 'Best Buys' tab lists out popular ingredients that will open up the option to make many more recipes.

If you have any feedback, I'd really appreciate it if you took the time to fill out my feedback survey [here.](https://docs.google.com/forms/d/e/1FAIpQLSfwoFzmMW6um5t2pzUnB7f0iB8C0iP2R_4OIiHqTUvzZG5W6A/viewform?usp=header)

[Software Demo Video](https://youtu.be/LZTql5tP-KA)

# Data Analysis Results

**Question 1: With a typical pantry, how many recipes can I make right now, and how much does the matching rule change that number?**

Each recipe is compared against the pantry one ingredient name at a time. Under strict matching, an ingredient counts as covered only when every word of its name appears in a single pantry item. Under loose matching, a pantry item also covers any longer ingredient that contains it, so "chicken" covers "chicken breast". I counted the recipes with nothing missing under both rules using my own pantry list. Loose matching finds more recipes, but it also lets through some bad matches, such as butter standing in for peanut butter. Strict matching is the safer default, and the loose option is there for people who want a longer list and are ok with a little manual work.

**Question 2: Which single ingredient would unlock the most new recipes if I bought it?**

This is the Best Buys tab. The program collects every recipe that is missing exactly one non-seasoning ingredient, then counts how often each missing ingredient appears. The ingredient with the highest count is the best purchase. The ranking showed me that a handful of common items account for a large share of the near misses.

Techniques used:

* **Filter:** title, must-contain, and cannot-contain filters, plus the pantry check itself.
* **Sort:** recipes are ordered by how many ingredients are missing, and every column in the results tables can be sorted by clicking its heading.
* **Aggregate:** counts of makeable recipes per tab and a frequency count of missing ingredients.
* **Data conversion:** the dataset stores lists as JSON text, so each row is parsed, lowercased, stripped of filler words such as "chopped", and reduced to singular forms so that "tomatoes" and "tomato" match.

# Development Environment


The program is in Python 3 and kept under version control with Git, hosted on GitHub. The interface uses tkinter, which ships with Python. To run it, install the libraries with `pip install pandas python-docx`, download `recipes_data.csv` from the Kaggle link above, and place it in the same folder as `pantry_pal.py`. Then start it with `python pantry_pal.py`. The program finds the CSV on its own. If the file is somewhere else, it asks for the location once and remembers it. Favorites and lists are saved to `library.json`, and the pantry is saved to `pantry.json`.

Libraries used:

* **pandas** to read the large CSV and drop incomplete rows
* **tkinter and ttk** for the window, tabs, tables, and pantry editor
* **threading and queue** so loading and searching happen in the background without freezing the window
* **python-docx** to export recipes as Word documents
* **json, re, functools, collections, and pathlib** from the standard library for parsing, text cleanup, caching, counting, saving data, and file paths

The full dataset needs several gigabytes of memory, so the program loads the first 200,000 recipes by default. That limit is a single constant near the top of the file.

# Useful Websites

* [Recipe Dataset (over 2M) on Kaggle](https://www.kaggle.com/datasets/wilmerarltstrmberg/recipe-dataset-over-2m)
* [pandas Documentation](https://pandas.pydata.org/docs/)
* [10 Minutes to pandas](https://pandas.pydata.org/docs/user_guide/10min.html)
* [Python tkinter Documentation](https://docs.python.org/3/library/tkinter.html)
* [Python tkinter.ttk Documentation](https://docs.python.org/3/library/tkinter.ttk.html)

# Future Work

* Recipe entry
* Quantity fetching
* Time needed for cooking