from tkinter import Tk, Label, Button, Entry, END
from tkinter import ttk
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os
from datetime import datetime

#  variable to store movie data
movie_data = []

# List of genres for the dropdown
genres = [
    "Drama", "Adventure", "Thriller", "Action", "Crime", "Comedy", "Mystery", "War", 
    "Fantasy", "Sci-Fi", "Biography", "Family", "Romance", "Animation", "History", 
    "Sport", "Western", "Music", "Horror", "Musical", "Film-Noir"
]

# Initialize the WebDriver (ChromeDriver in this case)
def create_driver():
    options = Options()
    #options.add_argument("--headless")  # headless mode (But not working with IMDB)
    driver = webdriver.Chrome(options=options) 
    return driver

# Scrape IMDb based on the user's input (Genre)
def get_movie_suggestions(genre):
    global movie_data
    driver = create_driver()
    driver.maximize_window() 
    try:
        # Go to IMDb's advanced search page
        driver.get("https://www.imdb.com/chart/top/?ref_=nv_mv_250")
        # To Decline the Cookies and Permissions
        try:
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//button[text()='Decline']")))
            cookies_decline = driver.find_element(By.XPATH, "//button[text()='Decline']").click()
        except:
            print("Loading.....")
        # To Open the Genre Filter
        genre_filter = driver.find_element(By.XPATH, "//*[@id='__next']/main/div/div[3]/section/div/div[2]/div/div[2]/div[2]/div/button").click()      
        All_genres_button = driver.find_element(By.XPATH, "//button[text()='Show all genres']").click() # To expand all_genres
        
        # Find genre button based on the user input
        genre_select = WebDriverWait(driver, 20).until(lambda x:x.find_element(By.XPATH, "//button[span[contains(text(), '"+genre+"')]]")).click()
        body_click = driver.find_element(By.XPATH, "/html/body/div[4]/div[2]/div/div[1]/button").click() # To remove the genre filter
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='__next']/main/div/div[3]/section/div/div[2]/div/ul"))) # Wait until all the elements are loaded
        
        # Get movie list
        Table = driver.find_element(By.XPATH, "//*[@id='__next']/main/div/div[3]/section/div/div[2]/div/ul") # Table with Movie data
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='__next']/main/div/div[3]/section/div/div[2]/div/ul/li/div[2]")))# Wait until all the elements are loaded
        movies = Table.find_elements(By.XPATH, "//*[@id='__next']/main/div/div[3]/section/div/div[2]/div/ul/li/div[2]") # To Find all the Movies
        movie_list = []
        
        for movie in movies:
            title = movie.text.split("\n")
            Movie_title = title[0].split(".", 1)[1].strip() if len(title) > 1 else "N/A" # Movie Title
            year = title[1].strip() if len(title) > 1 else "N/A"  # Movie Release Year
            runtime = title[2].strip() if len(title) > 2 else "N/A"  # Movie Runtime
            rating = title[4].strip() if len(title) > 4 else "N/A"  # Movie Rating
            votes = title[5].strip() if len(title) > 5 else "N/A"  # Movie Votes
            # If clause if rating is not available but votes are available
            if votes == "Rate":
                votes = rating
                rating = "N/A"
            movie_list.append([ Movie_title, year, runtime, rating, votes])
        # Update the movie_list to  movie_data variable
        movie_data = movie_list
        return movie_list

    except Exception as e:
        print(f"Error: {e}") # To handle Exceptions
        return []

    finally:
        driver.quit()

# Handle user input and display results in the GUI
def on_search_click():
    genre = genre_combobox.get()
    
    if not genre:
        # Clear previous data from treeview
        for row in treeview.get_children():
            treeview.delete(row)
        return
    
    # Call the scraping function
    suggestions = get_movie_suggestions(genre)
    
    # Clear previous data from treeview
    for row in treeview.get_children():
        treeview.delete(row)
    
    # Display results in the treeview
    if suggestions:
        for idx, suggestion in enumerate(suggestions, 1):
            treeview.insert("", "end", values=(idx, *suggestion))
    else:
        treeview.insert("", "end", values=("N/A", "No movies found.", "", "", "", ""))

# Function to export movie list to Excel
def export_to_excel():
    if not movie_data:
        print("No data to export")
        return

    # Generate a unique filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"output_{timestamp}.xlsx"
    
    # Create a pandas DataFrame
    df = pd.DataFrame(movie_data, columns=[ "Movie Title", "Year", "Runtime", "Rating", "Votes"])

    # Save the DataFrame to Excel
    df.index = df.index + 1 # To start index with 1
    df.to_excel(filename, index=True)
    print(f"Data exported to {filename}")

# Set up the Tkinter GUI
root = Tk()
root.title("IMDb Movie Suggestions")

# Instructions Label
instructions = """
1. Select the Genre from the Dropdown.
2. Click on Search (This will open the Chrome Driver and search for the input).
3. Once the Bot Fetches the data, it will be displayed in the white blank space in the form of a table.
4. You can click on Export to get this displayed data in Excel format, which will be downloaded in the same location where your Movie Suggestion App is located.
"""

label = Label(root, text=instructions, justify='left', padx=10, font=('Arial', 10, 'italic'))
label.grid(row=99, column=0, columnspan=2, pady=10)  # Placing at the bottom (row=99 or the last available row)

# Create Genre Label and Entry Box
genre_label = Label(root, text="Genre:")
genre_label.grid(row=0, column=0, padx=10, pady=5)
genre_entry = Entry(root, width=25)
genre_entry.grid(row=0, column=1, padx=10, pady=5)
genre_combobox = ttk.Combobox(root, values=genres, width=25)
genre_combobox.grid(row=0, column=1, padx=10, pady=5)
genre_combobox.set("Drama")  # Default selection
# Create Search Button
search_button = Button(root, text="Search", command=on_search_click)
search_button.grid(row=1, column=0, columnspan=2, pady=10)

# Create Export Button
export_button = Button(root, text="Export", command=export_to_excel)
export_button.grid(row=3, column=0, columnspan=2, pady=10)

# Create Treeview for displaying movie suggestions
treeview = ttk.Treeview(root, columns=("Index", "Movie Title", "Year", "Runtime", "Rating", "Votes"), show="headings", height=15)
treeview.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

# Define column headings
treeview.heading("Index", text="Index")
treeview.heading("Movie Title", text="Movie Title")
treeview.heading("Year", text="Year")
treeview.heading("Runtime", text="Runtime")
treeview.heading("Rating", text="Rating")
treeview.heading("Votes", text="Votes")

# Set column width
treeview.column("Index", width=50, anchor="center")
treeview.column("Movie Title", width=250, anchor="w")
treeview.column("Year", width=80, anchor="center")
treeview.column("Runtime", width=100, anchor="center")
treeview.column("Rating", width=80, anchor="center")
treeview.column("Votes", width=100, anchor="center")

# Start the Tkinter event loop
root.mainloop()
