import re
import os
import spacy
import pandas as pd
from spacy.matcher import Matcher
from pdfminer.high_level import extract_text as extract_text_from_pdf

directory_path = r"C:\\Users\\uzivatel\\Desktop\\CVs"

def get_contact_number(text):
    contact_number = None
    pattern = r"\+?\d{1,3}[-.\s]?\(?\d{1,4}\)?[-.\s]?\d{1,4}[-.\s]?\d{1,4}[-.\s]?\d{1,4}"
    match = re.search(pattern, text)
    if match:
        contact_number = match.group()
    return contact_number

def get_email(text):
    email = None
    pattern = r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b"
    match = re.search(pattern, text)
    if match:
        email = match.group()
    return email

def get_skills(text, skillset):
    found_skills = []
    lower_text = text.lower()
    for skill in skillset:
        pattern = r"\b{}\b".format(re.escape(skill.lower()))
        match = re.search(pattern, lower_text)
        if match:
            found_skills.append(skill)
    return found_skills

def get_name(resume_text):
    nlp = spacy.load('en_core_web_sm')
    matcher = Matcher(nlp.vocab)
    name_patterns = [
        [{'POS': 'PROPN'}, {'POS': 'PROPN'}],
        [{'POS': 'PROPN'}, {'POS': 'PROPN'}, {'POS': 'PROPN'}],
        [{'POS': 'PROPN'}, {'POS': 'PROPN'}, {'POS': 'PROPN'}, {'POS': 'PROPN'}]
    ]
    for pattern in name_patterns:
        matcher.add('NAME', [pattern])
    doc = nlp(resume_text)
    matches = matcher(doc)
    for match_id, start, end in matches:
        span = doc[start:end]
        return span.text
    return None

skillset = ["Python", "Data Analysis", "Analýza dat", "analytika", "datová analytika", "Machine Learning", "Strojové učení",
        "Communication", "Komunikace", "Project Management", "Projektové řízení", "Deep Learning", "Hluboké učení",
        "SQL", "Tableau", "Artificial Intelligence", "Umělá inteligence", "Data Mining", "Dolování dat",
        "Data Visualization", "Vizualizace dat", "Database Management", "Správa databází", "Big Data", "Velká data",
        "Statistical Analysis", "Statistická analýza", "Predictive Analytics", "Prediktivní analýza",
        "Business Intelligence", "Risk Management", "Řízení rizik", "Financial Analysis", "Finanční analýza",
        "Quantitative Analysis", "Kvantitativní analýza", "Budget Planning", "Plánování rozpočtu", "Cost Management",
        "Řízení nákladů", "Audit", "Compliance", "Dodržování předpisů", "Asset Management", "Správa aktiv",
        "Blockchain", "Cryptocurrency", "Kryptoměna", "Forex Trading", "Forexové obchodování",
        "Investment Strategies", "Investiční strategie", "Portfolio Management", "Správa portfolia",
        "Equity Trading", "Obchodování s akciemi", "Market Analysis", "Analýza trhu", "ERP Systems", "ERP systémy",
        "SAP", "Oracle", "Cloud Computing", "AWS", "Azure", "DevOps", "Agile Methodologies", "Agilní metodiky",
        "Scrum", "Cybersecurity", "Kybernetická bezpečnost", "Penetration Testing", "Penetrační testování",
        "IT Support", "IT podpora", "Software Development", "Vývoj softwaru", "Web Development", "Vývoj webu",
        "JavaScript", "Java", "C#", "C++", "Ruby", "PHP", "Swift", "Kotlin", "Node.js", "React", "Angular",
        "UX/UI Design", "Graphic Design", "Grafický design", "Digital Marketing", "Digitální marketing", "SEO",
        "Content Marketing", "Marketing obsahu", "Social Media Management", "Správa sociálních médií",
        "Email Marketing", "Emailový marketing", "PPC Campaigns", "PPC kampaně", "Customer Relationship Management",
        "Řízení vztahů se zákazníky", "Sales Strategies", "Prodejní strategie", "Negotiation", "Vyjednávání",
        "Customer Service", "Zákaznický servis", "Public Speaking", "Veřejné vystupování", "Team Leadership", "Vedení týmu",
        "Human Resources", "Lidské zdroje", "Recruiting", "Nábor", "Performance Management", "Řízení výkonu",
        "Employee Training", "Školení zaměstnanců", "Payroll Management", "Správa mezd", "Time Management",
        "Řízení času", "Crisis Management", "Krízové řízení", "Strategic Planning", "Strategické plánování",
        "Corporate Finance", "Firemní finance", "Mergers & Acquisitions", "Fúze a akvizice", "Corporate Law",
        "Obchodní právo", "Taxation", "Daně", "E-commerce", "Supply Chain Management", "Řízení dodavatelského řetězce",
        "Logistics", "Logistika", "Product Management", "Řízení produktů", "Quality Assurance", "Zajištění kvality",
        "ISO Standards", "ISO normy", "Environmental Compliance", "Dodržování environmentálních předpisů",
        "Health and Safety Management", "Řízení zdraví a bezpečnosti", "Sustainability", "Udržitelnost",
        "Corporate Social Responsibility", "Společenská odpovědnost firem", "Event Planning", "Plánování akcí",
        "Workshop Facilitation", "Vedení workshopů", "Coaching", "Koučink", "Mentoring", "Mentoring", "Accounting",
        "Účetnictví", "Financial Reporting", "Finanční reportování", "Regulatory Compliance", "Regulační dodržování",
        "Credit Risk", "Úvěrové riziko", "Market Risk", "Tržní riziko", "Operational Risk", "Operační riziko",
        "Treasury Management", "Řízení pokladny", "Derivatives", "Deriváty", "Capital Markets", "Kapitálové trhy",
        "Trading", "Obchodování", "Business Development", "Obchodní rozvoj", "Stakeholder Management",
        "Řízení zainteresovaných stran", "Leadership", "Vedení", "Decision Making", "Rozhodování", "Problem Solving",
        "Řešení problémů", "Process Improvement", "Zlepšování procesů", "Change Management", "Řízení změn",
        "Innovation", "Inovace", "Client Management", "Řízení klientů", "Vendor Management", "Řízení dodavatelů",
        "Contract Negotiation", "Vyjednávání smluv", "Financial Modeling", "Finanční modelování", "Strategy Development",
        "Rozvoj strategie", "Presentation Skills", "Prezentační dovednosti", "Analytical Skills", "Analytické dovednosti",
        "Reporting", "Reportování", "Data Warehousing", "Skladování dat", "Algorithms", "Algoritmy", "Power BI", "SPSS",
        "MS Office", "Powerpoint", "Excel", "Teams", "Access", "Outlook", "SAS", "Keboola", "Czech", "English", "German",
        "French", "čeština", "angličtina", "němčina", "český jazyk", "německý jazyk", "francouzský jazyk", "anglický jazyk",
        "francouzština", "BigQuery", "GoodData", "Snowflake", "Programming", "ETL", "Git", "Hadoop", "Grafana", "JIRA",
        "Airflow", "Big data"
]
    
extracted_data = []
for file in os.listdir(directory_path):
        if file.endswith('.pdf'):
            file_path = os.path.join(directory_path, file)
            text_content = extract_text_from_pdf(file_path)
            
            candidate_name = get_name(text_content) or "Name not found"
            phone_number = get_contact_number(text_content) or "Contact Number not found"
            email_address = get_email(text_content) or "Email not found"
            skills = get_skills(text_content, skillset) or ["No skills found"]

            extracted_data.append({
                'File Name': file,
                'Candidate Name': candidate_name,
                'Phone': phone_number,
                'Mail': email_address,
                'Skills': ', '.join(skills)
            })

df = pd.DataFrame(extracted_data)
df.to_excel(f'{directory_path}\\extracted_resume_data.xlsx', index=False)
print(f"Data successfully extracted and saved to '{directory_path}'.")
