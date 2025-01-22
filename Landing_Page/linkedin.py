import csv
from bs4 import BeautifulSoup

# Path to the offline LinkedIn profile HTML file
html_file = "linkedin_profile.html"

def scrape_linkedin(html_file):
    try:
        # Open the HTML file
        with open(html_file, 'r', encoding='utf-8') as file:
            soup = BeautifulSoup(file, 'html.parser')

        # Extract job titles
        job_titles = [job.text.strip() for job in soup.select(".job-title")]

        # Extract companies
        companies = [company.text.strip() for company in soup.select(".company-name")]

        # Extract industries
        industries = [industry.text.strip() for industry in soup.select(".industry")]

        # Combine extracted data into a structured format
        data = [{"Job Title": title, "Company": company, "Industry": industry} 
                for title, company, industry in zip(job_titles, companies, industries)]

        # Save data to a CSV file
        with open("linkedin_data.csv", "w", newline="", encoding="utf-8") as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=["Job Title", "Company", "Industry"])
            writer.writeheader()
            writer.writerows(data)

        print("Scraping completed. Data saved to linkedin_data.csv.")

    except Exception as e:
        print(f"Error occurred: {e}")

# Run the scraper
scrape_linkedin(html_file)
