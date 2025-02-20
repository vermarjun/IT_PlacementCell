from playwright.sync_api import sync_playwright
import csv

# DIRECTIONS TO USE:

# STEP 1.) Close Chrome before

# STEP 2.) OPEN CMD, Paste this and press enter: "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222
# NOTE: In above command replace the path with chrome's exe file in your system

# Step3) Edit below file paths and variables according to need and jsut run the script

csv_file_path = "./data.csv" # ./NameOfFile.extension Enter the path of csv file that has data
closed_caption = "alokkushwaha@ggu.ac.in"
website_link = "https://email.gov.in"
brochure_file_path = "C:/Users/verma/Downloads/Placement Brochure 2024 - 2025.pdf"
name_row = 1 # In csv file where name field is (starting from 0)
email_row = 2 # In csv file where emial field is (starting from 0)

with sync_playwright() as playwright:
    browser = playwright.chromium.connect_over_cdp("http://localhost:9222")

    default_context = browser.contexts[0]

    page = default_context.pages[0]

    with open(csv_file_path, mode='r', newline='', encoding='utf-8') as file:
        reader = csv.reader(file)
        print(reader)
        for row in reader:
            if len(row) >= 4:  # Ensure the row has enough columns
                name = row[name_row].strip()  
                email = row[email_row].strip() 
                print(f"Mailing to => Name: {name}, Email: {email}")

                page.goto(website_link)

                newMessagebtn =  page.locator("div[id=zb__NEW_MENU]")
                newMessagebtn.click()

                to =  page.locator("input[id=zv__COMPOSE-1_to_control]")
                to.fill(email)
                cc =  page.locator("input[id=zv__COMPOSE-1_cc_control]")
                cc.fill(closed_caption)
                subject =  page.locator("input[id=zv__COMPOSE-1_subject_control]")
                subject.fill("Invitation for Placement Drive in School of Studies in Engineering & Technology, Guru Ghasidas University, A Central University, Bilaspur")

                attach_button_selector = '#zb__COMPOSE-1___attachments_btn'
                page.click(attach_button_selector)
                file_input_selector = 'input[type="file"]'  
                page.set_input_files(file_input_selector, brochure_file_path)
                
                iframe_selector = 'iframe[id="ZmHtmlEditor1_body_ifr"]' 
                page.wait_for_selector(iframe_selector)

                editor_iframe = page.frame_locator(iframe_selector)
                
                contenteditable_selector = 'body#tinymce' 
                editor_iframe.locator(contenteditable_selector).fill(f'Dear {name},\n\nGreetings from the Placement Cell at the School of Studies in Engineering & Technology, Guru Ghasidas Vishwavidyalaya (A Central University), Bilaspur, C.G.\n\nI hope you are doing well.\n\nOur school, established in 1997, admits students to the B.Tech program through the JEE Mains and M.Tech through the GATE Examination. Many of our former students have found placements in diverse MNCs and successful startups. Our University is rated NAAC A++ and among the top 50 Indian Universities in the Times World Ranking.\n\nAs part of our continuous efforts to provide the finest opportunities to our students and help pave their path to a successful career, we would be delighted if you consider the possibility of conducting your recruitment drive for the B.Tech Batch of 2025. We have students from the following branches: CSE / IT / ECE / CHEM / MECH / IPE / CIVIL and we can assure you of providing high-quality candidates from these disciplines. Additionally, if your company offers summer internships, you are invited to participate in our internship drive for the Batch of 2026.\n\nWe have a rich legacy of producing highly skilled and motivated engineers who could be valuable assets to your company. Our goal is to assist our students in building a successful career with your organization.\n\nWe are open to both physical and virtual recruitment drives according to your preferences and convenience. Please write back to us with your tentative drive dates, and feel free to request more information or specify any requirements from our end.\n\nWe are excited to hear back from you. I\'ve attached our placement brochure to this email.\n\nThanks & Warm Regards,\nDr Alok Kumar Singh Kushwaha,\nPlacement Co-ordinator School of Studies in Engineering & Technology,\nGuru Ghasidas Vishwavidyalaya (A Central University), Bilaspur\nEmail - alokkushwaha@ggu.ac.in\nEmail - tpo@ggu.ac.in')
                
                upload_complete_selector = 'a.AttLink'  
                page.wait_for_selector(upload_complete_selector, timeout=100000) 
                
                sendBtn =  page.locator("div[id=zb__COMPOSE-1__SEND_MENU]")
                sendBtn.click()
