import os
from pptx import Presentation
from pptx.util import Inches

# Create a presentation object
filename = 'Python_Data_Structures.pptx'

if os.path.exists(filename):
    prs = Presentation(filename)
else:
    prs = Presentation()

def add_slide(prs, slide_layout, title_text, content_text):
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]

    title.text = title_text
    content.text = content_text

# Title Slide
title_slide_layout = prs.slide_layouts[0]  # 0 is the layout for title slide
slide = prs.slides.add_slide(title_slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]

title.text = "Introduction to Python Data Structures"
subtitle.text = "Strings, Lists, Tuples, and Dictionaries"

# Strings Section
content_slide_layout = prs.slide_layouts[1]  # 1 is the layout for title and content
strings_slides = [
    ("What is a String?", """
    - A string is a sequence of characters.
    - Strings are used to represent text.
    - In Python, strings are immutable, meaning they cannot be changed after creation.
    - Strings can be enclosed in single quotes (' '), or double quotes (" ").
    """),
    ("Creating Strings", """
    - Single quotes: 'Hello'
    - Double quotes: "Hello"
    """),
    ("Accessing Characters", """
    - Zero-based indexing
    - Accessing the first character: my_string[0]
    - Accessing the last character: my_string[-1]
    """),
    ("Slicing Strings", """
    - Access a substring: my_string[1:4]
    - Slicing with a step: my_string[0:5:2]
    """),
    ("String Methods", """
    - Convert to uppercase: my_string.upper()
    - Convert to lowercase: my_string.lower()
    - Split a string: my_string.split()
    - Join a list of strings: ' '.join(list_of_strings)
    """),
    ("String Concatenation", """
    - Using + operator: 'Hello' + ' ' + 'World'
    - Using join method: ' '.join(['Hello', 'World'])
    """),
    ("String Formatting", """
    - Using str.format(): 'Hello {}'.format('World')
    - Using f-strings (Python 3.6+): f'Hello {name}'
    - Limiting a float to n decimal places: '{:.2f}'.format(3.20159)
    """),
    ("Common String Operations", """
    - Find a substring: my_string.find('substring')
    - Replace a substring: my_string.replace('old', 'new')
    - Check if alphanumeric: my_string.isalnum()
    - Strip whitespace: my_string.strip()
    """)
]

for title_text, content_text in strings_slides:
    add_slide(prs, content_slide_layout, title_text, content_text)

# Lists Section
lists_slides = [
    ("What is a List?", """
    - A list is an ordered collection of items.
    - Lists are mutable, meaning they can be modified.
    - Lists can contain elements of different data types.
    - Created using square brackets: [1, 2, 3]
    """),
    ("Creating a List", """
    - An empty list: empty_list = []
    - A list of integers: int_list = [1, 2, 3, 4, 5]
    - A list of strings: str_list = ["apple", "banana", "cherry"]
    - A mixed list: mixed_list = [1, "apple", 3.20, True]
    """),
    ("Accessing Elements", """
    - Zero-based indexing
    - Accessing the first element: my_list[0]
    - Accessing the last element: my_list[-1]
    """),
    ("Modifying Elements", """
    - Modify an element by assigning a new value: my_list[1] = "blueberry"
    """),
    ("Adding Elements", """
    - Append: my_list.append("cherry")
    - Insert: my_list.insert(1, "blueberry")
    - Extend: my_list.extend(["date", "elderberry"])
    """),
    ("Removing Elements", """
    - Remove by value: my_list.remove("banana")
    - Remove by index and return: my_list.pop(0)
    - Delete by index: del my_list[0]
    """),
    ("Slicing", """
    - Access a subset of a list: my_list[1:3]
    - Slicing with a step: my_list[0:4:2]
    """),
    ("List Comprehensions", """
    - Create lists concisely
    - Example: squares = [x**2 for x in range(5)]
    """),
    ("Common List Methods", """
    - len(list)
    - list.sort()
    - list.reverse()
    - list.index(value)
    - list.count(value)
    """)
]

for title_text, content_text in lists_slides:
    add_slide(prs, content_slide_layout, title_text, content_text)

# Tuples Section
tuples_slides = [
    ("What is a Tuple?", """
    - A tuple is an ordered collection of items.
    - Tuples are immutable, meaning they cannot be modified after creation.
    - Tuples can contain elements of different data types.
    - Created using parentheses: (1, 2, 3)
    """),
    ("Creating a Tuple", """
    - An empty tuple: empty_tuple = ()
    - A tuple of integers: int_tuple = (1, 2, 3, 4, 5)
    - A tuple of strings: str_tuple = ("apple", "banana", "cherry")
    - A mixed tuple: mixed_tuple = (1, "apple", 3.14, True)
    """),
    ("Accessing Elements in a Tuple", """
    - Zero-based indexing
    - Accessing the first element: my_tuple[0]
    - Accessing the last element: my_tuple[-1]
    """),
    ("Common Tuple Operations", """
    - len(tuple)
    - tuple.index(value)
    - tuple.count(value)
    """)
]

for title_text, content_text in tuples_slides:
    add_slide(prs, content_slide_layout, title_text, content_text)

# Dictionaries Section
dictionaries_slides = [
    ("What is a Dictionary?", """
    - A dictionary is an unordered collection of key-value pairs.
    - Keys must be unique and immutable (e.g., strings, numbers, tuples).
    - Values can be of any data type and can be duplicated.
    - Created using curly braces: {'key1': 'value1', 'key2': 'value2'}
    """),
    ("Creating a Dictionary", """
    - An empty dictionary: empty_dict = {}
    - A dictionary with key-value pairs: my_dict = {'name': 'Alice', 'age': 25}
    """),
    ("Accessing Values", """
    - Access value by key: my_dict['name']
    - Using get() method: my_dict.get('name')
    """),
    ("Modifying Values", """
    - Change an existing value: my_dict['age'] = 26
    - Add a new key-value pair: my_dict['city'] = 'New York'
    """),
    ("Removing Key-Value Pairs", """
    - Using pop() method: my_dict.pop('age')
    - Using del statement: del my_dict['name']
    - Using popitem() method to remove the last inserted item: my_dict.popitem()
    """),
    ("Dictionary Methods", """
    - keys(): Returns a view object of all keys
    - values(): Returns a view object of all values
    - items(): Returns a view object of all key-value pairs
    - update(): Updates the dictionary with elements from another dictionary
    """),
    ("Looping Through a Dictionary", """
    - Loop through keys: for key in my_dict
    - Loop through values: for value in my_dict.values()
    - Loop through key-value pairs: for key, value in my_dict.items()
    """),
    ("Dictionary Comprehensions", """
    - Create dictionaries concisely
    - Example: squares = {x: x**2 for x in range(5)}
    """),
    ("Nested Dictionaries", """
    - A dictionary within a dictionary
    - Example: nested_dict = {'parent': {'child': 'value'}}
    """),
    ("Differences Between Lists and Dictionaries", """
    - Lists are ordered collections, while dictionaries are unordered.
    - Lists use integer indices to access elements, while dictionaries use keys.
    - Lists are mutable and allow duplicate elements, while dictionaries require unique keys.
    - Use lists for ordered data and dictionaries for key-value pairs.
    """)
]

for title_text, content_text in dictionaries_slides:
    add_slide(prs, content_slide_layout, title_text, content_text)

# Save the presentation
prs.save('Python_Data_Structures.pptx')