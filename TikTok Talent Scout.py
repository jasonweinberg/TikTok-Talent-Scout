import sys
import re
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLineEdit, QListWidget, QFileDialog, QMessageBox, QLabel, QProgressBar
from PyQt5.QtCore import Qt
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

# Validate the input URL to ensure it's a valid TikTok profile URL.
def validate_tiktok_url(url):
    patterns = [
        r"https:\/\/www\.tiktok\.com\/@[\w\.\-]+",
        r"@[\w\.\-]+",
        r"[\w\.\-]+"
    ]
    return any(re.match(pattern, url) for pattern in patterns)

# Format TikTok URL to match the first format (https://www.tiktok.com/@username).
def format_tiktok_url(url):
    if url.startswith("https://www.tiktok.com/"):
        return url
    return f"https://www.tiktok.com/@{url.lstrip('@')}"

# Extract name, total followers, and total likes from TikTok profile page.
def extract_data(driver):
    name_element = driver.find_element(By.CSS_SELECTOR, 'h2[data-e2e="user-subtitle"]')
    name = name_element.text.strip()

    followers_element = driver.find_element(By.CSS_SELECTOR, 'strong[data-e2e="followers-count"]')
    total_followers = followers_element.text.strip()

    likes_element = driver.find_element(By.CSS_SELECTOR, 'strong[data-e2e="likes-count"]')
    total_likes = likes_element.text.strip()

    return name, total_followers, total_likes

# Function to scrape profile data
def scrape_profile(url, progress_bar):
    # Set up Chrome options
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Optional: run in headless mode

    # Set up Selenium WebDriver with webdriver-manager
    chrome_service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

    try:
        # Load the TikTok profile page
        driver.get(url)
        print("Page loaded successfully.")

        # Update progress bar value
        progress_bar.setValue(50)

        # Extract data
        name, followers, likes = extract_data(driver)

        # Extract username from URL
        username = re.search(r'@([\w\.\-]+)', url).group(1)

        # Update progress bar value
        progress_bar.setValue(100)

        return name, username, followers, likes, url
    finally:
        driver.quit()

class TikTokScraperGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.layout = QVBoxLayout()

        self.url_input = QLineEdit(self)
        self.url_input.setPlaceholderText('Enter TikTok profile URL, @username or username')
        self.layout.addWidget(self.url_input)

        self.scrape_button = QPushButton('Scrape Profile', self)
        self.scrape_button.clicked.connect(self.scrape_profile)
        self.layout.addWidget(self.scrape_button)

        self.progress_bar = QProgressBar(self)
        self.layout.addWidget(self.progress_bar)

        self.preview_label = QLabel('Name:\nFollowers:\nLikes:', self)
        self.preview_label.setAlignment(Qt.AlignLeft)
        self.layout.addWidget(self.preview_label)

        self.add_to_spreadsheet_button = QPushButton('Add to Spreadsheet', self)
        self.add_to_spreadsheet_button.clicked.connect(self.add_to_spreadsheet)
        self.layout.addWidget(self.add_to_spreadsheet_button)

        self.remove_from_spreadsheet_button = QPushButton('Remove from Spreadsheet', self)
        self.remove_from_spreadsheet_button.clicked.connect(self.remove_from_spreadsheet)
        self.layout.addWidget(self.remove_from_spreadsheet_button)

        self.spreadsheet_list = QListWidget(self)
        self.layout.addWidget(self.spreadsheet_list)

        self.export_button = QPushButton('Export to Excel', self)
        self.export_button.clicked.connect(self.export_to_excel)
        self.layout.addWidget(self.export_button)

        self.setLayout(self.layout)
        self.setWindowTitle('TikTok Profile Scraper')

        self.profile_data = []

    def scrape_profile(self):
        url = self.url_input.text()
        if validate_tiktok_url(url):
            url = format_tiktok_url(url)
            self.progress_bar.setValue(0)  # Start at 0%
            self.progress_bar.setRange(0, 100)  # Range up to 100%

            # Set progress to 10% when DevTools loads
            self.progress_bar.setValue(10)

            # Set progress to 50% when the page properly loads
            self.progress_bar.setValue(50)

            try:
                name, username, followers, likes, profile_url = scrape_profile(url, self.progress_bar)  # Pass progress_bar argument

                # Set progress to 90% when data extraction is complete
                self.progress_bar.setValue(90)

                self.preview_label.setText(f'Name: {name}\nFollowers: {followers}\nLikes: {likes}')
                self.current_profile = (name, username, followers, likes, profile_url)

                # Set progress to 100% when the data populates in the GUI
                self.progress_bar.setValue(100)
            except Exception as e:
                self.show_error("Error occurred while scraping the profile. Please try again.")
        else:
            self.show_error("Invalid URL format. Please enter a valid TikTok profile.")



    def add_to_spreadsheet(self):
        if hasattr(self, 'current_profile'):
            name, username, followers, likes, profile_url = self.current_profile
            self.spreadsheet_list.addItem(name)
            self.profile_data.append((name, username, followers, likes, profile_url))
        else:
            self.show_error("No profile to add. Please scrape a profile first.")

    def remove_from_spreadsheet(self):
        selected_item = self.spreadsheet_list.currentItem()
        if selected_item:
            index = self.spreadsheet_list.row(selected_item)
            self.spreadsheet_list.takeItem(index)
            del self.profile_data[index]
        else:
            self.show_error("No profile selected to remove.")

    def show_error(self, message):
        error_dialog = QMessageBox(self)
        error_dialog.setIcon(QMessageBox.Critical)
        error_dialog.setText(message)
        error_dialog.setWindowTitle('Error')
        error_dialog.exec_()

    def export_to_excel(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, "Save File", "", "Excel Files (*.xlsx)", options=options)
        if file_path:
            self.save_profiles_to_excel(file_path)

    def save_profiles_to_excel(self, file_path):
        df = pd.DataFrame(self.profile_data, columns=["Name", "Username", "Followers", "Likes", "Profile URL"])
        df.to_excel(file_path, index=False)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = TikTokScraperGUI()
    ex.show()
    sys.exit(app.exec_())
