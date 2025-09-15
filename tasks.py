import sqlite3
import os
import csv
import datetime
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.Email.ImapSmtp import ImapSmtp
from robocorp.tasks import task
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
from dotenv import load_dotenv

#load env variables
load_dotenv()

# --- Configuration ---
DB_NAME = os.getenv("DB_NAME")
EXCEL_FILE = os.getenv("EXCEL_FILE")
CSV_FILE = os.getenv("CSV_FILE")

EMAIL_FROM = os.getenv("EMAIL_FROM")
EMAIL_TO = os.getenv("EMAIL_TO")
EMAIL_PASSWORD = os.getenv("EMAIL_APP_PASSWORD")

browser = Selenium()


# --- Database Setup ---
def setup_db():
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS movies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            movie_name TEXT,
            tomatometer_score TEXT,
            audience_score TEXT,
            storyline TEXT,
            rating TEXT,
            genres TEXT,
            review_1 TEXT,
            review_2 TEXT,
            review_3 TEXT,
            review_4 TEXT,
            review_5 TEXT,
            status TEXT
        )
    """)
    conn.commit()
    conn.close()
    print("Database setup completed.")


def insert_result(movie, data):
    conn = sqlite3.connect(DB_NAME)
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO movies (movie_name, tomatometer_score, audience_score, storyline, rating, genres,
                            review_1, review_2, review_3, review_4, review_5, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (movie, *data))
    conn.commit()
    conn.close()
    print(f"Inserted data for movie: {movie}")


def handle_cookie_banner():
    """Handle cookie banner if present"""
    try:
        if browser.does_page_contain_element("id:onetrust-accept-btn-handler"):
            browser.click_button("id:onetrust-accept-btn-handler")
            print("Cookie banner accepted.")
    except Exception:
        pass


def extract_movie_data():
    """
    Extracts movie details from a TMDB movie page:
    - TOMATOMETER (NA)
    - Audience/User Score
    - Storyline/Overview
    - Rating (PG-13, etc.)
    - Genres
    - Top 5 reviews (text only)
    """
    tomatometer = "NA"
    audience_score = ""
    storyline = ""
    rating = ""
    genres = ""
    reviews = [""] * 5
    status = "failure"

    try:
        # --- Audience Score ---
        try:
            score_elem = browser.find_element("//div[contains(@class,'user_score_chart')]")
            audience_score = browser.get_element_attribute(score_elem, "data-percent")
            if not audience_score:
                audience_score = ""
        except:
            audience_score = ""

        # --- Storyline ---
        try:
            storyline_elem = browser.find_element("css:div.overview > p")
            storyline = browser.get_text(storyline_elem).strip()
        except:
            storyline = ""

        # --- Rating ---
        try:
            rating_elem = browser.find_element("css:span.certification")
            rating = browser.get_text(rating_elem).strip()
        except:
            rating = ""

        # --- Genres ---
        try:
            genre_elems = browser.find_elements("//span[contains(@class,'genres')]/a")
            genres = ", ".join([browser.get_text(g).strip() for g in genre_elems])
        except:
            genres = ""

        # --- Reviews ---
        try:
            current_url = browser.get_location()
            reviews_url = current_url + "/reviews"
            browser.go_to(reviews_url)

            # Wait for reviews to load
            try:
                WebDriverWait(browser.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div.review_container'))
                )
            except:
                pass

            # Get review containers using the working method from friend's code
            containers = browser.get_webelements("css:div.review_container > div.content")
            texts = []
            
            for container in containers[:5]:
                try:
                    # Find the <p> inside the nested path relative to the container
                    paragraphs = browser.find_elements("css:div.teaser > p", parent=container)
                    if paragraphs:
                        review_text = browser.get_text(paragraphs[0]).strip()
                        texts.append(review_text)
                    else:
                        texts.append("")
                except:
                    texts.append("")
                    
            # Pad to 5 reviews
            while len(texts) < 5:
                texts.append("")
            reviews = texts
            
        except Exception as e:
            print(f"Error extracting reviews: {e}")
            reviews = [""] * 5

        # --- Status ---
        if storyline or rating or genres or audience_score:
            status = "success"

    except Exception as e:
        print(f"Error extracting movie data: {e}")
        status = "failure"

    return (tomatometer, audience_score, storyline, rating, genres, *reviews, status)


def search_movie(movie):
    print(f"\nProcessing: {movie}")
    
    try:
        # Handle cookie banner first
        handle_cookie_banner()
        
        search_url = f"https://www.themoviedb.org/search?query={movie.replace(' ', '%20')}"
        browser.go_to(search_url)

        try:
            # Wait for search results with increased timeout
            WebDriverWait(browser.driver, 15).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.card.v4.tight'))
            )
        except:
            print("Search results not loaded")
            insert_result(movie, ("NA", "", "", "", "", "", "", "", "", "", "failure"))
            return

        cards = browser.find_elements("css:div.card.v4.tight")
        exact_matches = []
        print(f"Found {len(cards)} results for '{movie}'")

        for i, card in enumerate(cards):
            try:
                # Get title
                title_elem = browser.find_element("css:div.details div.title h2", parent=card)
                title = browser.get_text(title_elem).strip()

                # Ensure it's a movie (not TV show or person)
                link_elem = browser.find_element("css:div.details a", parent=card)
                href = browser.get_element_attribute(link_elem, "href")
                if "/movie/" not in href:
                    continue

                # Get release date for comparison
                try:
                    release_elem = browser.find_element("css:span.release_date", parent=card)
                    release_text = browser.get_text(release_elem).strip()
                    try:
                        release_date = datetime.datetime.strptime(release_text, "%B %d, %Y")
                    except:
                        release_date = datetime.datetime.min
                except:
                    release_date = datetime.datetime.min

                # Check for exact match
                if title.lower() == movie.lower():
                    exact_matches.append((title, release_date, title_elem))
                    print(f"Found exact match: '{title}' ({release_date.date() if release_date != datetime.datetime.min else 'Unknown date'})")
                    
            except Exception as e:
                print(f"Error processing card {i}: {e}")
                continue

        if not exact_matches:
            print(f"No exact match found for '{movie}'")
            insert_result(movie, ("NA", "", "", "", "", "", "", "", "", "", "failure"))
            return

        # Select the most recent exact match
        most_recent = max(exact_matches, key=lambda x: x[1])
        print(f"Selected most recent: '{most_recent[0]}' ({most_recent[1].date() if most_recent[1] != datetime.datetime.min else 'Unknown date'})")

        # Click on the selected movie
        browser.click_element(most_recent[2])

        # Extract movie data
        data = extract_movie_data()
        insert_result(movie, data)
        
    except Exception as e:
        print(f"Error in search_movie for '{movie}': {e}")
        insert_result(movie, ("NA", "", "", "", "", "", "", "", "", "", f"Processing error: {str(e)}"))


def read_excel_and_process():
    excel = Files()
    excel.open_workbook(EXCEL_FILE)
    sheet = excel.read_worksheet_as_table(header=True)
    movies = []
    
    for row in sheet:
        # Try multiple possible column names
        movie = (row.get("Movies") or 
                row.get("Movie Name") or 
                row.get("Movie") or 
                row.get("Title") or
                row.get("movies"))  # Added lowercase option
        if movie and movie.strip():
            movies.append(movie.strip())
    excel.close_workbook()
    
    print(f"Found {len(movies)} movies to process: {movies}")

    # Initialize browser once
    browser.open_available_browser("https://www.themoviedb.org", headless=False)
    browser.set_selenium_timeout(60)
    browser.set_window_size(1920, 1080)

    for movie in movies:
        try:
            search_movie(movie)
        except Exception as e:
            print(f"Error processing '{movie}': {e}")
            insert_result(movie, ("NA", "", "", "", "", "", "", "", "", "", f"Processing error: {str(e)}"))


# --- Export CSV ---
def export_to_csv():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM movies")
    data = cursor.fetchall()
    cursor.execute("PRAGMA table_info(movies)")
    columns = [row[1] for row in cursor.fetchall()]
    conn.close()

    with open(CSV_FILE, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(columns)
        writer.writerows(data)
    print(f"CSV exported to {CSV_FILE}")


# --- Send Email ---
def send_email():
    export_to_csv()
    try:
        mail = ImapSmtp(smtp_server="smtp.gmail.com", smtp_port=587)
        mail.authorize(EMAIL_FROM, EMAIL_PASSWORD)
        mail.send_message(
            sender=EMAIL_FROM,
            recipients=EMAIL_TO,
            subject="TMDB Movie Data Extraction",
            body="See attached CSV file with movie data.",
            attachments=[CSV_FILE] if os.path.exists(CSV_FILE) else None
        )
        print("Email sent successfully!")
    except Exception as e:
        print(f"Error sending email: {e}")


# --- Main Task ---
@task
def main():
    if not os.path.exists(EXCEL_FILE):
        print(f"Excel file {EXCEL_FILE} not found!")
        return
    try:
        setup_db()
        read_excel_and_process()
        send_email()
        print("Movie extraction completed!")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        try:
            browser.close_all_browsers()
        except:
            pass


if __name__ == "__main__":
    main()