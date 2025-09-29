This Python Script contains official email ID and password of INSTRUO so don't share it to anyone outside.

Prerequisites: Open Powershell and type this command :-  pip install pandas openpyxl




This is a one time process for each domain.

1. Copy the prompt and paste in ChatGPT "change the following content type to [your domain (say robotics)]. Make sure the {name} doesn't gets affected because that is being used in automated Python script. Here is the email content

Dear {name},

We are delighted to invite {name} to collaborate with INSTRUO 14, the annual techno-management festival of IIEST Shibpur, scheduled from October 31st to November 2nd, 2025.

INSTRUO is among Eastern India’s largest technical festivals, attracting over 10,000 students from premier institutes in Kolkata, along with top participants from IITs, NITs, and other leading universities across the country. It serves as a dynamic platform for innovation, knowledge exchange, and industry–academia partnerships.

With your expertise and leadership in your industry, {name} perfectly aligns with the spirit of technological and educational excellence that we champion. Our flagship events, spanning technology, mobility, energy, education, media, and innovation, are among the festival’s biggest highlights, offering a unique opportunity for your brand to engage directly with talented young engineers, innovators, and future leaders.

We would be glad to explore collaboration avenues that match your vision, and would love to schedule a brief call to discuss the possibilities further.

Warm regards,
Aayush Sarkar
Sponsorship Associate
Team INSTRUO
Contact: 7367999558
"

2. Paste the generated email content back in email_body variable in send_emails.py script. Make sure that you don't affect the inverted commas ("").
3. Copy only the company name and email addresses from the spreadsheet and paste it in contacts.xlsx file. (DO NOT CHANGE THE Name AND Email COLUMN NAME AS IT WOULD RESULT IN ERROR.)
4. Right click in the folder and click open terminal then type python send_email.py


Sit back and cazzzzzz!!!
