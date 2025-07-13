from playwright.sync_api import Playwright, sync_playwright, expect
import os
import pandas as pd
from time import sleep, perf_counter
from datetime import datetime
import yagmail
import schedule

def send_email_report(file_path, to_email):
    yag = yagmail.SMTP(user="taneshka02mehta04@gmail.com", password="vbhctvzirjcgsshr")
    subject = "🌤️🗞️ Daily News Report"
    body = "Hi,\n\nPlease find attached today's news report.\n\nRegards,\nNewsBot"
    yag.send(to=to_email, subject=subject, contents=body, attachments=file_path)
    print(f"✅ Email sent to {to_email}")

dtt = datetime.now() 
f_dtt = dtt.strftime('%d-%m-%Y')

url = 'https://www.thehindu.com/'

headlines = []
newss = []
date_nz = []
ss = []

def run(playwright: Playwright) -> None:
    # Browser [Chrome, Firefox, WebKit]
    # browser = playwright.firefox.lanch(headless=False)

    browser = playwright.chromium.launch(headless=False)
    
    # Browser Screen Size
    # context = browser.new_context(viewport={ 'width': 2560, 'height': 1440 })
    context = browser.new_context()
    page = context.new_page()
    page.goto(url)
    sleep(1)

    # Aiming for Latest News   
    page.locator('//h2[@class="title-patch "]').click()
    sleep(2)

    # Total count of latest news present on UI
    count_nz = page.locator('//ul[@class="timeline-with-img"]//li').count()
    print(count_nz)
    nz = int(count_nz)
    print(nz)

    for y in range(nz):

        # News Link Extraction
        li = page.locator('//ul[@class="timeline-with-img"]//li//h3/a').nth(y).get_attribute('href')
        newss.append(li)
        sleep(0.5)

        # Headlines Extraction
        head = page.locator('//ul[@class="timeline-with-img"]//li//h3').nth(y).inner_text()
        print(head)
        headlines.append(head)
        sleep(0.5)

        # Date of Published News 
        date = page.locator('//ul[@class="timeline-with-img"]//li//div[@class="news-time time"]').nth(y).inner_text()
        print(date)
        date_nz.append(date)
        sleep(0.5)

        # Source for Published News
        source = page.locator('//ul[@class="timeline-with-img"]//li//div[@class="author-name"]//a[contains(@class, "person-name")]').nth(y).inner_text()
        print(source)
        ss.append(source)
        sleep(0.5)

    my_dic = {
        'Headlines': headlines,
        'News': newss,
        'Published News Date': date_nz,
        'Source': ss
        }
    df1 = pd.DataFrame(data=my_dic)
    files_path = '/Users/taneshkamehta/Documents/RPA/Web_Automation_Notifications/Daily_News_Report/ '

    # Get the absolute path of the Excel file
    absolute_path = os.path.abspath(files_path)

    # Extract the directory portion of the absolute path
    directory_path = os.path.dirname(absolute_path)
    # Specify the file name for the Excel file
    file_name = 'News_Report_%s.xlsx' %f_dtt
    # Combine the folder path and file name to get the full path
    file_path = os.path.join(directory_path, file_name)
    # Save the DataFrame to the specified location (full path)
    df1.to_excel(file_path, sheet_name='error', index=False)

    # Send email after saving Excel
    send_email_report(file_path, to_email="taneshka.mehta@gmail.com")

    context.close()
    browser.close()

def fn_run():
    with sync_playwright() as playwright:
        run(playwright)

if __name__ == "__main__":
    start = perf_counter()
    fn_run()
    end = perf_counter()
    print(f'\n---------------\n Finished in {round(end-start, 2)} second(s)')

    # # Schedule to run daily at 08:00 AM
    # # schedule.every().day.at("10:00").do(fn_run)
    # # print("⏰ Scheduler started. Waiting for 08:00 AM daily job...")

    # schedule.every(1).minutes.do(fn_run)
    # print("⏰ Scheduler started. Waiting for 1 Minute job...")


    # while True:
    #     schedule.run_pending()
    #     sleep(60)
