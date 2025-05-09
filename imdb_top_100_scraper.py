import json
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
from webdriver_manager.chrome import ChromeDriverManager
import re

IMDB_URL = "https://www.imdb.com/chart/top/"
EXCEL_FILENAME = "IMDb_Top_100.xlsx"
JSON_FILENAME = "IMDb_Top_100.json"

def setup_driver():
    options = Options()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

def extract_movies(driver):
    movie_divs = driver.find_elements(By.XPATH, "//div[contains(@class, 'cli-parent')]")
    movies = []
    for i, movie_div in enumerate(movie_divs):
        if i >= 100:
            break
        try:
            movie_data = {
                'title': re.sub(r'^\d+\.\s*', '', movie_div.find_element(By.XPATH, ".//h3[@class='ipc-title__text']").text),
                'year': movie_div.find_elements(By.XPATH, ".//span[contains(@class, 'cli-title-metadata-item')]")[0].text,
                'duration': movie_div.find_elements(By.XPATH, ".//span[contains(@class, 'cli-title-metadata-item')]")[1].text,
                'content_rating': movie_div.find_elements(By.XPATH, ".//span[contains(@class, 'cli-title-metadata-item')]")[2].text,
                'audience_rating': movie_div.find_element(By.XPATH, "..//span[@class='ipc-rating-star--rating']").text
            }
            movies.append(movie_data)
        except Exception as e:
            print(f"Error extracting movie: {e}")
    return movies

def save_to_excel(movies, filename):
    df = pd.DataFrame(movies)
    df.to_excel(filename, index=False)

    wb = load_workbook(filename)
    ws = wb.active

    for column_cells in ws.columns:
        max_length = max((len(str(cell.value)) for cell in column_cells), default=0)
        adjusted_width = max_length + 6
        column_letter = column_cells[0].column_letter
        ws.column_dimensions[column_letter].width = adjusted_width

    wb.save(filename)

def save_to_json(movies, filename):
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(movies, f, ensure_ascii=False, indent=4)


def print_movies(movies):
    for movie in movies:
        print(f"Title: {movie['title']}")
        print(f"Year: {movie['year']}")
        print(f"Duration: {movie['duration']}")
        print(f"Content Rating: {movie['content_rating']}")
        print(f"Audience Rating: {movie['audience_rating']}")
        print("-" * 150)

def main():
    driver = setup_driver()
    driver.get(IMDB_URL)
    try:
        movies = extract_movies(driver)

        if not movies:
            print("No movies extracted.")
            return
        
        save_to_excel(movies, EXCEL_FILENAME)
        save_to_json(movies, JSON_FILENAME)
        print_movies(movies)
        print(f"\nSaved movies to '{EXCEL_FILENAME}' and '{JSON_FILENAME}'")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()