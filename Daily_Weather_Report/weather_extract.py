from playwright.sync_api import Playwright, sync_playwright, expect
import os
import pandas as pd
from time import sleep, perf_counter
from datetime import datetime
import yagmail
import schedule

def send_email_report(file_path, to_email):
    yag = yagmail.SMTP(user="taneshka02mehta04@gmail.com", password="rygwrjxlqjcmwpuj")
    subject = "🌤️ Daily Weather Extraction Report"
    body = "Hi,\n\nPlease find attached today's weather report.\n\nRegards,\nWeatherBot"
    yag.send(to=to_email, subject=subject, contents=body, attachments=file_path)
    print(f"✅ Email sent to {to_email}")

dtt = datetime.now() 
f_dtt = dtt.strftime('%d-%m-%Y')

url = 'https://weather.com/'

xls_file = 'city_data.xlsx'
df = pd.read_excel(xls_file)
citys = []
timee = []
curr_loc = []
temval = []
condval = []
dayhi = []
nightlo = []

ct = df['City'].values.tolist()

def run(playwright: Playwright) -> None:
    # Browser [Chrome, Firefox, WebKit]
    # browser = playwright.firefox.lanch(headless=False)

    browser = playwright.chromium.launch(headless=False)
    
    # Browser Screen Size
    # context = browser.new_context(viewport={ 'width': 2560, 'height': 1440 })
    context = browser.new_context()
    page = context.new_page()
    page.goto(url)
    sleep(0.5)

    for idx, n in enumerate(ct):
        sleep(1)
        page.locator('(//input[contains(@placeholder, "Search City")])[1]').fill(ct[idx])
        page.locator('(//div[contains(@id, "listbox")]//button)[1]').click()
        sleep(2)
        curr_time = page.locator('//span[contains(@class, "timestamp")]').inner_text()
        # print(curr_time)
        timee.append(curr_time)
        sleep(0.2)
        ex_loc = page.locator('//div[contains(@class, "CurrentConditions")]//h1').inner_text()
        # print(ex_loc)
        curr_loc.append(ex_loc)
        sleep(0.2)
        temp = page.locator('//div[contains(@class, "CurrentConditions")]//span[contains(@class, "tempValue")]').inner_text()
        # print(temp)
        temval.append(temp)
        sleep(0.2)
        cond = page.locator('//div[contains(@class, "CurrentConditions--phrase")]').inner_text()
        # print(cond)
        condval.append(cond)
        sleep(0.2)
        day = page.locator('(//div[contains(@class, "tempHiLoValue")]//span[@data-testid="TemperatureValue"])[1]').inner_text()
        # print(day)
        dayhi.append(day)
        sleep(0.2)
        night = page.locator('(//div[contains(@class, "tempHiLoValue")]//span[@data-testid="TemperatureValue"])[2]').inner_text()
        # print(night)
        nightlo.append(night)
        sleep(1)

    my_dic = {
        'Current Location': curr_loc,
        'Time': timee,
        'Temprature (°C)': temval,
        'Condition': condval,
        'High/Day Temperature (°C)': dayhi,
        'Low/Night Temperature (°C)': nightlo
        }
    df1 = pd.DataFrame(data=my_dic)
    excel_path = xls_file

    # Get the absolute path of the Excel file
    absolute_path = os.path.abspath(excel_path)

    # Extract the directory portion of the absolute path
    directory_path = os.path.dirname(absolute_path)
    # Specify the file name for the Excel file
    file_name = 'Weather_Extractions_%s.xlsx' %f_dtt
    # Combine the folder path and file name to get the full path
    file_path = os.path.join(directory_path, file_name)
    # Save the DataFrame to the specified location (full path)
    df1.to_excel(file_path, sheet_name='error', index=False)

    # Send email after saving Excel
    send_email_report(file_path, to_email="taneshka.mehta@gmail.com")

    context.close()
    browser.close()


def fn_run():
    # information('function f')
    with sync_playwright() as playwright:
        run(playwright)

if __name__ == "__main__":
    start = perf_counter()
    fn_run()
    end = perf_counter()
    print(f'\n---------------\n Finished in {round(end-start, 2)} second(s)')

    # Schedule to run daily at 08:00 AM
    # schedule.every().day.at("10:00").do(fn_run)
    # print("⏰ Scheduler started. Waiting for 08:00 AM daily job...")

    schedule.every(1).minutes.do(fn_run)
    print("⏰ Scheduler started. Waiting for 1 Minute job...")


    while True:
        schedule.run_pending()
        sleep(60)
