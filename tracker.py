#!/usr/bin/env python3
"""
Job Search Email Tracker & Pipeline Manager
============================================
Универсальный инструмент для автоматического отслеживания поиска работы по почтовым ящикам.
Рассчитан на регулярные запуски в будущем:
- Поддерживает любое количество аккаунтов Gmail (хранятся в accounts.json).
- Парсит письма начиная с указанной даты (флаг --since или по умолчанию 01-Aug-2026).
- Интеллектуально распознает компании, позиции и стадии рекрутинга по универсальным паттернам (любые ATS, job-борды, прямая переписка).
- Автоматически связывает письма одного работодателя, даже если отклик и ответ пришли на разные ящики или с разных адресов.
- Отсекает информационный шум (сервисные письма, рассылки вакансий, рекламу).
- Рассчитывает количество дней без ответа (для follow-up).
- Генерирует CSV-таблицу (Excel/Numbers/Sheets) и интерактивный веб-дашборд.
"""

import os
import sys
import getpass
import json
import re
import csv
import argparse
import imaplib
import email
import socket
from email.header import decode_header
from datetime import datetime
from collections import defaultdict
from html import escape, unescape

# Защита от бесконечного зависания сетевых соединений IMAP
socket.setdefaulttimeout(20)

# По умолчанию дата начала текущего поиска работы
DEFAULT_SINCE_DATE = "01-Aug-2026"

# Известные платформы ATS (Applicant Tracking Systems)
KNOWN_ATS_DOMAINS = [
    "greenhouse", "greenhouse-mail", "lever.co", "ashbyhq", "smartrecruiters",
    "workday", "workdayjobs", "myworkday", "workable", "traffit", "jobvite",
    "applytojob", "successfactors", "jobs2web", "erecruiter", "zohorecruit",
    "nofluffjobs", "recruitee", "teamtailor", "pinpointhq", "bamboohr", "breezy",
    "linkedin", "avature"
]

# Контекст вакансии: письмо должно его содержать, чтобы считаться откликом/ответом
JOB_CONTEXT_RE = re.compile(
    r"your application|applying|applied|application for|candida|\bposition\b|\brole\b|"
    r"hiring process|recruitment|job search|ваканси|кандидат|резюме|отклик|"
    r"aplikacj|rekrutac|stanowisk|\bcv\b",
    re.IGNORECASE
)

# Пользовательские правила (заполняются из ignore.txt / overrides.json в main)
USER_IGNORE_PATTERNS = []
USER_STATUS_OVERRIDES = {}

# Бесплатные почтовые сервисы (для них не извлекаем имя компании из домена)
COMMON_FREE_MAIL_DOMAINS = [
    "gmail.com", "googlemail.com", "yahoo.com", "outlook.com", "hotmail.com",
    "mail.ru", "yandex.ru", "icloud.com", "me.com", "proton.me", "protonmail.com"
]

# Категории отправителей, которые гарантированно не относятся к откликам
NON_JOB_DOMAINS = [
    "accounts.google.com", "googleplay-noreply@google.com", "analytics-noreply@google.com",
    "googlestore-noreply@google.com", "notifications@github.com", "noreply@github.com",
    "notices.dropbox.com", "feedback.decathlon.com", "paczkomaty.pl", "uaudio.com",
    "heavyocity.com", "wavelet-audio.com", "spotify.com", "support@hirify.me",
    "hey.simplify.jobs", "support@resumeworded.com", "laravel-news.com", "asana.com",
    "cloudflare.com", "santanderconsumer.pl", "messaging.logitech.com", "pkobp.pl",
    "mbank.pl", "santander.pl", "pekao.com.pl", "ing.pl", "aliorbank.pl", "millennium.pl"
]

# Первичные маркеры поиска работы
JOB_KEYWORDS = [
    # Ru
    "ваканси", "отклик", "резюме", "собеседован", "интервью", "тестовое", "оффер", 
    "приглашени", "работодател", "hh.ru", "хабр карьер", "headhunter", "djinni",
    # En
    "application", "job", "career", "candidate", "interview", "resume", "cv",
    "applied", "role", "position", "offer", "recruiting", "talent", "assessment",
    # Pl
    "aplikacj", "stanowisk", "rekrutacj"
]

# Паттерны для отказа
REJECTION_PATTERNS = [
    r"unfortunately",
    r"regret to inform",
    r"decided to pursue other candidates",
    r"moving forward with other candidates",
    r"not moving forward",
    r"will not be moving forward",
    r"not to move forward",
    r"won'?t be moving forward",
    r"other candidates who(se)?",
    r"decided not to proceed",
    r"position has been filled",
    r"wish you (the best|success) in your job search",
    r"at this time, we have chosen",
    r"we have chosen to move forward with",
    r"not a (good )?fit at this time",
    r"already shortlisted candidates whose profiles more closely align",
    r"shortlisted candidates whose profiles",
    r"can'?t expand (it )?with a new position",
    r"stay with the team that I have",
    r"pity it didn'?t work out",
    r"decided not to move forward",
    r"won'?t be able to offer",
    r"not to proceed with your application",
    r"closed this position",
    r"position is no longer available",
    r"unfortunate position",
    r"decided to (?:move|progress|proceed|continue|go) (?:forward )?with (?:another|other) candidates?",
    r"decided to progress with candidates",
    r"(?:closer|more closely) (?:align|match)",
    r"not (?:be )?(?:able to )?(?:progress|proceed|move forward) with your (?:application|candidacy)",
    r"will not be progressing",
    r"not been selected",
    r"we regret",
    # Ru
    r"к сожалению",
    r"не готовы пригласить",
    r"отклонен",
    r"отказ",
    r"вынуждены отказать",
    r"остановили свой выбор на другом",
    r"остановились на другом кандидате",
    r"решили продолжить с другим кандидатом",
    r"вернуться к вашей кандидатуре позже",
    r"в резерв",
    r"не готовы продолжить общение",
    r"не можем предложить",
    r"вакансия закрыта",
    r"отклонил вашу кандидатуру",
    # Pl
    r"nie możemy zaproponować",
    r"zdecydowaliśmy się na innego",
    r"nie przechodzisz do kolejnego etapu"
]

# Паттерны для интервью / тестирования
INTERVIEW_PATTERNS = [
    r"invitation to interview",
    r"schedule a (call|chat|time|meeting|screening)",
    r"screening call",
    r"phone screen",
    r"technical interview",
    r"calendly\.com",
    r"online assessment stage",
    r"passed the .*? hiring assessment",
    r"congratulations! you passed",
    r"complete the online assessment",
    r"take home (task|assignment)",
    r"interview invitation",
    # Ru
    r"приглашаем на собеседование",
    r"приглашение на собеседование",
    r"пригласить вас на (собеседование|интервью|созвон)",
    r"хотели бы пообщаться",
    r"тестовое задание",
    r"онлайн-встреч",
    # Pl
    r"zaproszenie na rozmowę",
    r"chcielibyśmy zaprosić na rozmowę"
]

# Паттерны для оффера
OFFER_PATTERNS = [
    r"formal offer",
    r"job offer letter",
    r"pleased to offer you",
    r"delighted to offer you",
    r"рады предложить вам оффер",
    r"предложение о работе"
]

# Паттерны для подтверждения получения отклика
CONFIRMATION_PATTERNS = [
    r"thank you for (your )?applying",
    r"thank you for your application",
    r"application (has been )?received",
    r"application (was |has been )?submitted",
    r"application has been sent",
    r"we(?:['’]ve| have)? received your (?:application|resume|cv)",
    r"you(?:['’]ve| have) applied",
    r"thank you for your interest in",
    r"thanks for applying",
    r"your application was sent to",
    r"your application was viewed by",
    r"your application arrived safely",
    r"application for .*? was submitted successfully",
    # Ru
    r"ваша заявка принята",
    r"спасибо за отклик",
    r"ваш отклик принят",
    r"ваш отклик доставлен",
    r"вы откликнулись на вакансию",
    r"отклик на вакансию.*отправлен",
    # Pl
    r"dziękujemy za (?:przesłanie|twoją) (?:cv|aplikacj)",
    r"informacja o twojej aplikacji",
    r"twoja aplikacja została przesłana",
    r"we(?:['’]ve| have) got your application"
]

def decode_mime_words(raw_header):
    """Декодирует MIME заголовок в строку UTF-8."""
    if not raw_header:
        return ""
    decoded_fragments = []
    try:
        parts = decode_header(raw_header)
        for fragment, encoding in parts:
            if isinstance(fragment, bytes):
                try:
                    encoding = encoding or 'utf-8'
                    decoded_fragments.append(fragment.decode(encoding, errors='replace'))
                except (LookupError, UnicodeDecodeError):
                    decoded_fragments.append(fragment.decode('latin1', errors='replace'))
            else:
                decoded_fragments.append(str(fragment))
    except Exception:
        return str(raw_header)
    return "".join(decoded_fragments).strip()

def clean_subject_for_grouping(subject: str) -> str:
    """Удаляет префиксы ответов и пересылок (Re, Fwd, Odp, На)."""
    cleaned = re.sub(r'^\s*(?:(?:re|fwd|fw|odp|на)\s*:\s*)+', '', subject, flags=re.IGNORECASE).strip()
    return cleaned

def extract_email_address(from_header: str) -> str:
    """Извлекает email из поля заголовка."""
    match = re.search(r'<([^>]+)>', from_header)
    if match:
        return match.group(1).lower().strip()
    match = re.search(r'[\w\.-]+@[\w\.-]+', from_header)
    if match:
        return match.group(0).lower().strip()
    return from_header.lower().strip()

def extract_sender_name(from_header: str) -> str:
    """Извлекает отображаемое имя контакта."""
    decoded = decode_mime_words(from_header)
    match = re.match(r'^"?([^"<]+)"?\s*<', decoded)
    if match:
        return match.group(1).strip()
    return decoded.split('<')[0].replace('"', '').strip()

def extract_domain(email_str: str) -> str:
    """Извлекает чистый домен из адреса."""
    m = re.search(r'@([a-zA-Z0-9\.\-]+)', email_str)
    return m.group(1).lower() if m else ''

def get_body_text(msg) -> str:
    """Извлекает текстовое содержимое письма с очисткой HTML."""
    text_content = []
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition", ""))
            if "attachment" in content_disposition:
                continue
            if content_type in ["text/plain", "text/html"]:
                try:
                    payload = part.get_payload(decode=True)
                    if payload:
                        charset = part.get_content_charset() or "utf-8"
                        try:
                            text = payload.decode(charset, errors="replace")
                        except (LookupError, UnicodeDecodeError):
                            text = payload.decode("latin1", errors="replace")
                        if content_type == "text/html":
                            text = re.sub(r'<style.*?</style>', '', text, flags=re.DOTALL | re.IGNORECASE)
                            text = re.sub(r'<script.*?</script>', '', text, flags=re.DOTALL | re.IGNORECASE)
                            text = re.sub(r'<[^>]+>', ' ', text)
                            text = re.sub(r'\s+', ' ', text)
                        text_content.append(text)
                except Exception:
                    pass
    else:
        try:
            payload = msg.get_payload(decode=True)
            if payload:
                charset = msg.get_content_charset() or "utf-8"
                try:
                    text = payload.decode(charset, errors="replace")
                except (LookupError, UnicodeDecodeError):
                    text = payload.decode("latin1", errors="replace")
                if msg.get_content_type() == "text/html":
                    text = re.sub(r'<style.*?</style>', '', text, flags=re.DOTALL | re.IGNORECASE)
                    text = re.sub(r'<script.*?</script>', '', text, flags=re.DOTALL | re.IGNORECASE)
                    text = re.sub(r'<[^>]+>', ' ', text)
                    text = re.sub(r'\s+', ' ', text)
                text_content.append(text)
        except Exception:
            pass

    return " ".join(text_content).strip()

def is_noise_or_marketing(subject: str, from_hdr: str, body: str) -> bool:
    """
    Универсальный фильтр нерелевантных писем:
    - Сервисные уведомления и безопасность
    - Подписки / рассылки вакансий (Job Alerts, digests)
    - Создание аккаунтов без отклика
    - Банковские, транзакционные и рекламные письма
    """
    from_email = extract_email_address(from_hdr).lower()
    subj_lower = subject.lower()
    body_snippet = body[:800].lower()

    # 0. Пользовательский список исключений (ignore.txt)
    haystack = f"{from_hdr} {subject}".lower()
    if any(p in haystack for p in USER_IGNORE_PATTERNS):
        return True

    # 0b. Корпоративные рассылки вакансий (jobs2web, "New jobs posted", job alerts)
    if re.search(r"jobalert|jobnotification|job-alert|talentcommunity", from_email) or \
       re.search(r"new jobs posted|job alert|jobs? for you|new opportunities for you", subj_lower):
        return True

    # 0c. Сервисные уведомления Google о доступе к данным / входе через Google
    if "shared some google account data" in subj_lower or "noreply-accounts@google.com" in from_email:
        return True

    # 1. Известные нерелевантные сервисы
    for domain in NON_JOB_DOMAINS:
        if domain in from_email or domain in from_hdr.lower():
            return True

    # 2. Безопасность аккаунтов (Google, Apple, MS, GitHub)
    if any(domain in from_email for domain in ["accounts.google.com", "google.com", "appleid.apple.com", "github.com"]):
        if any(w in subj_lower or w in body_snippet for w in [
            "парол", "password", "безопасност", "security alert", "вход в аккаунт", 
            "sign-in", "verification code", "код подтверждения", "двухэтапн", "2-step",
            "oauth application", "run failed: ci", "ci workflow"
        ]):
            return True

    # 3. LinkedIn рассылки и алерты
    if any(addr in from_email for addr in [
        "jobalerts-noreply@linkedin.com", "newsletters-noreply@linkedin.com",
        "updates-noreply@linkedin.com", "groups-noreply@linkedin.com",
        "career-interests-noreply@linkedin.com"
    ]):
        return True

    if "jobs-noreply@linkedin.com" in from_email:
        # Пропускаем только подтверждения отправленных откликов
        is_real_app = any(pat in subj_lower for pat in [
            "your application was sent to", "you applied to", "your application to", "your application was viewed by"
        ])
        if not is_real_app:
            return True

    if "messaging-digest-noreply@linkedin.com" in from_email:
        if not any(k in subj_lower for k in ["interview", "собеседован", "ваканси"]):
            return True

    # 4. Шаблоны рассылок вакансий (Job Alerts, подборки)
    job_alert_patterns = [
        r"\b\d+\s+(?:new\s+jobs?|новых?\s+ваканси)",
        r"(?:jobs?\s+for\s+you|jobs?\s+matching|job\s+alert|recommended\s+jobs?)",
        r"(?:подборка\s+вакансий|вакансии\s+дня|подходящие\s+вакансии|новые\s+вакансии\s+для\s+вас)",
        r"(?:новые\s+вакансии\s+по\s+запросу|дайджест\s+вакансий|рекомендуемые\s+вакансии)",
        r"(?:top\s+job\s+picks|jobs\s+you\s+may\s+be\s+interested\s+in|fresh\s+jobs)",
        r"(?:weekly\s+job\s+alert|daily\s+job\s+alert|see\s+more\s+jobs|is\s+hiring\s+for\s+a\s+remote)",
        r"(?:expand\s+your\s+search|jobs\s+that\s+match\s+your\s+profile|looking\s+for\s+a\s+new\s+job\?)"
    ]
    if any(re.search(pat, subj_lower) for pat in job_alert_patterns):
        return True

    # 5. Регистрация и активация профилей без конкретного отклика
    account_creation_patterns = [
        r"twoje konto na .* jest już aktywne",
        r"welcome - thanks for creating",
        r"careers account created",
        r"thank you for creating an account",
        r"verify your candidate account",
        r"candidate profile - account details",
        r"welcome to aina",
        r"witamy na dou polska",
        r"аккаунт успешно создан",
        r"подтвердите ваш email",
        r"verify your email address"
    ]
    if any(re.search(pat, subj_lower) for pat in account_creation_patterns):
        return True

    # 6. Банковские, транзакционные и служебные уведомления
    if any(w in subj_lower for w in [
        "podziel się z nami opinią", "wizyty w sklepie", "terms are changing", 
        "najważniejsza lekcja to", "deal days", "mobywatel", "prawdziwą aplikację", 
        "stronę logowania", "aktywacja aplikacji"
    ]):
        return True

    return False

def domain_to_company(domain: str) -> str:
    """Извлекает имя бренда из домена корпоративной почты."""
    if not domain or any(free in domain for free in COMMON_FREE_MAIL_DOMAINS):
        return ""
    if any(ats in domain for ats in KNOWN_ATS_DOMAINS):
        return ""
    
    parts = domain.split(".")
    ignore_parts = {
        "com", "org", "net", "pl", "eu", "io", "co", "de", "uk", "me", "ai", "app",
        "jobs", "careers", "career", "mail", "email", "notifications", "notification",
        "hr", "hire", "candidate", "candidates", "recruit", "recruitment", "wd1", "us"
    }
    clean = [p for p in parts if p not in ignore_parts]
    if clean:
        base = clean[-1]
        base = re.sub(r'(inc|group|tech|technologies|recruitment|recruiting)$', '', base, flags=re.I)
        if len(base) >= 2:
            return base.capitalize()
    return ""

def normalize_company_name(name: str) -> str:
    """Универсальная нормализация названия компании для связывания писем."""
    if not name or name in ["Не определена", "Unknown"]:
        return ""
    clean = name.lower()
    clean = re.sub(r'["«»\'`“”]', '', clean)
    clean = re.sub(r'\.(?:com|pl|io|eu|ai|co|net|org|store)\b', '', clean)
    suffixes = [
        r'\b(?:llc|inc|corp|corporation|ltd|limited|gmbh|sa|bv|plc|sp\.\s*z\s*o\.o\.|sp\s*z\s*o\s*o)\b',
        r'\b(?:ооо|зао|оао|пао|ао|нко)\b',
        r'\b(?:technologies|technology|tech|software|solutions|services|group|labs|team|recruiting|recruitment|talent|interactive|systems|consulting)\b'
    ]
    for s in suffixes:
        clean = re.sub(s, '', clean, flags=re.IGNORECASE)
    clean = re.sub(r'[^a-zA-Zа-яА-Я0-9]', '', clean)
    return clean.strip()

TECH_TITLE_RE = re.compile(
    r'\b(?:(?:Lead|Senior|Principal|Staff|Head of|Director(?: of)?|VP(?: of)?|Chief|Junior|Mid|Middle)?\s*'
    r'(?:Software|Data|Cloud|DevOps|Platform|Site Reliability|Backend|Frontend|Full[- ]?stack|Mobile|Security|Product|Engineering|QA|Test|Architecture|Enterprise|Solutions?|IT)?\s*'
    r'(?:Engineer|Developer|Architect|Manager|Lead|Officer|Director|Specialist|Consultant|CTO|Team Leader|Engineering Manager|Cooperation))\b',
    re.IGNORECASE
)

def clean_company_name(name: str) -> str:
    """Очищает название компании от HR-суффиксов и юридических форм."""
    if not name:
        return ""
    name = re.sub(r'[\'"`“”«»]', '', name).strip()
    name = re.sub(r'\s*(?:Recruiting|Recruitment|Team|Hiring|Careers|Jobs|HR|Talent|Talent Acquisition|Global Talent Acquisition|Talent Attraction and Acquisition|Group|via\s+.*|System|no-reply|noreply|notifications?).*$', '', name, flags=re.I).strip()
    name = re.sub(r"['’]s$", "", name, flags=re.I).strip()
    name = re.sub(r'\s+(?:Sp\.\s*z\s*o\.o\.|Sp\s*z\s*o\s*o|S\.A\.|SA|LLC|Inc\.?|Corp\.?|Ltd\.?|GmbH)$', '', name, flags=re.I).strip()
    return name.strip()

def clean_position(pos: str, company: str = "") -> str:
    """Универсальная очистка и валидация наименования должности."""
    if not pos:
        return ""
    pos = unescape(pos).strip()
    pos = re.sub(r'[\'"`“”«»]', '', pos).strip()
    pos = re.sub(r'^(?:the|a|an)\s+', '', pos, flags=re.I).strip()
    # Удаляем идентификаторы заявок (R1543646 - , 1711773 )
    pos = re.sub(r'^[A-Z0-9]{5,12}\s*[-–—:]\s*', '', pos).strip()
    # Удаляем суффикс компании: "at Company", "with Company", "@ Company"
    pos = re.sub(r'\s+(?:at|with|here\s+at|in|w|we|@)\s+[A-Za-z0-9\s&/\\._-]+$', '', pos, flags=re.I).strip()
    if company and company.lower() in pos.lower():
        pos = re.sub(rf'\b(?:at|with|in|for)?\s*{re.escape(company)}\b', '', pos, flags=re.I).strip()
    # Удаляем висящие союзы и предлоги в конце строки ("and", "or", "for", "at", "with", "in")
    pos = re.sub(r'\s+(?:and|or|for|at|with|in)\s*$', '', pos, flags=re.I).strip()
    pos = pos.strip(' .,!?:;-–—')
    if len(pos) < 3 or len(pos) > 130:
        return ""
    bad_phrases = [
        "thank you", "thanks for", "update", "feedback", "your application", "quick check-in",
        "job search", "recruitment", "a job can be exciting", "interest in"
    ]
    if any(b in pos.lower() for b in bad_phrases):
        return ""
    return pos

def parse_metadata_from_email(subject: str, from_hdr: str, body: str):
    """
    Универсальное извлечение компании и позиции.
    Работает с любыми ATS, площадками и прямыми ответами рекрутеров.
    """
    company = ""
    position = ""

    from_name = extract_sender_name(from_hdr)
    from_email = extract_email_address(from_hdr)
    domain = extract_domain(from_hdr)

    # ---------------- 1. ИЗВЛЕЧЕНИЕ КОМПАНИИ ----------------
    # 1. LinkedIn: "your application was sent to <Company>"
    m_li = re.search(r'application (?:was )?(?:sent to|viewed by)\s+([^–—\-|,\.!]+)', subject, re.I)
    if m_li:
        company = clean_company_name(m_li.group(1).strip())

    # 1b. NoFluffJobs: в теле письма "Company: <Company>"
    if not company:
        m_nfj_comp = re.search(r'Company:\s*([A-Za-z0-9\s&]{2,30}?)(?:\s+Your application|\n|$)', body, re.I)
        if m_nfj_comp:
            company = clean_company_name(m_nfj_comp.group(1))

    # 2. Внутри кавычек From: ("HR@Bayer.com" <system@successfactors.eu>)
    if not company:
        m_inner = re.search(r'\"([a-zA-Z0-9\.\-_]+)@([a-zA-Z0-9\-_]+)\.[a-zA-Z]{2,}\"', from_hdr)
        if m_inner:
            cand_dom = m_inner.group(2).lower()
            if cand_dom not in ["system", "gmail", "google"]:
                company = cand_dom.capitalize()

    # 3. Workday username: iqvia@myworkday.com -> IQVIA
    if not company:
        m_workday = re.match(r'^([a-zA-Z0-9\-]+)@(?:[a-zA-Z0-9\-]+\.)?myworkday(?:jobs)?\.com', from_email)
        if m_workday:
            company = m_workday.group(1).upper()

    # 4. Домен компании в username ATS: recruitment.profitroom.com@viazohorecruit.eu
    if not company:
        m_user_dom = re.search(r'([a-zA-Z0-9\-]+)\.(?:com|pl|eu|io)@', from_email)
        if m_user_dom:
            cand_u = m_user_dom.group(1).capitalize()
            if cand_u.lower() not in ["google", "gmail"]:
                company = cand_u

    # 5. Прямой корпоративный домен отправителя (приоритет над именем рекрутера!)
    if not company and domain:
        dom_comp = domain_to_company(domain)
        if dom_comp:
            company = dom_comp

    # 6. Префикс в теме письма: "Company - Informacja o Twojej aplikacji", "sun.store | we've got..."
    if not company:
        m_pref = re.match(r'^([A-Za-z0-9\.\s&]{2,30}?)\s*[-–—|]\s*(?:we[’\']ve got|informacja|dziękujemy|your application|feedback|new job|thanks)', subject, re.I)
        if m_pref:
            company = clean_company_name(m_pref.group(1))

    # 7. Имя отправителя From Name (если не платформа и не ATS)
    if not company and from_name:
        cand_from = clean_company_name(from_name)
        if cand_from and cand_from.lower() not in ['google', 'gmail', 'support', 'notification', 'notifications', 'workable', 'traffit', 'headhunter', 'hh.ru', 'linkedin', 'no fluff jobs', 'hr system', 'staffing board']:
            # Если в теме есть "at <Company>", это может быть точнее имени рекрутера
            m_at = re.search(r'\b(?:at|@|with|to)\s+([A-Za-z0-9\s&]{2,35}?)(?:!|\.|,|-|—|–|\s*\(|\s*\[|\s*$)', subject, re.I)
            if m_at:
                cand_at = clean_company_name(m_at.group(1))
                if cand_at.lower() not in ["the", "a", "our", "present", "this stage"] and not any(k in cand_at.lower() for k in ["role", "position"]):
                    company = cand_at
            if not company:
                company = cand_from

    # 8. Суффикс "at / @ / with / to <Company>" в теме
    if not company:
        m_at = re.search(r'\b(?:at|@|with|to)\s+([A-Za-z0-9\s&]{2,35}?)(?:!|\.|,|-|—|–|\s*\(|\s*\[|\s*$)', subject, re.I)
        if m_at:
            cand_at = clean_company_name(m_at.group(1))
            if cand_at.lower() not in ["the", "a", "our", "present", "this stage"] and not any(k in cand_at.lower() for k in ["role", "position"]):
                company = cand_at

    # ---------------- 2. ИЗВЛЕЧЕНИЕ ПОЗИЦИИ ----------------
    body_clean = unescape(body[:1500])
    body_one_line = re.sub(r'[ \t]+', ' ', body_clean)

    # 1. LinkedIn Easy Apply confirmation pattern
    if company and company != "Не определена":
        m_li_exact = re.search(rf'Your application was (?:sent to|viewed by)\s+{re.escape(company)}\s+([A-Za-z0-9\s\-_/,\(\)&|]+?)\s+{re.escape(company)}', body_one_line, re.I)
        if m_li_exact:
            cand = clean_position(m_li_exact.group(1), company)
            if cand:
                position = cand
    if not position:
        m_li_pos = re.search(r'Your application was (?:sent to|viewed by)\s+(?:[A-Za-z0-9\s&]+?)\s+([A-Za-z0-9\s\-_/,\(\)&|]+?)\s+(?:[A-Za-z0-9\s&]+?)\s+(?:Poland|Warsaw|Krakow|Remote|View job)\b', body_one_line, re.I)
        if m_li_pos:
            cand = clean_position(m_li_pos.group(1), company)
            if cand:
                position = cand

    # 2. NoFluffJobs confirmation
    if not position:
        m_nfj = re.search(r'Job offer:\s*([^\n\r]+?)(?:\s+Company:|\n|$)', body_clean, re.I)
        if m_nfj:
            position = clean_position(m_nfj.group(1), company)

    # 3. В теме письма: "application for <Position>", "applied to the <Position> role"
    if not position:
        m_subj_pos = re.search(r'(?:application for|applied to the|applied for|interest in the|regarding your application for the)\s+(?:the\s+)?([A-Za-z0-9\s\-_/,\(\)&|]+?)(?:\s+(?:role|position|job)\b|\s+[@|at|w]\s+|\.|\n|$)', subject, re.I)
        if m_subj_pos:
            position = clean_position(m_subj_pos.group(1), company)

    # 4. В теме письма: "<Position> at <Company>"
    if not position:
        m_subj_at = re.search(r'^([A-Za-z0-9\s\-_/,\(\)&|]{3,80}?)\s+(?:at|@)\s+([A-Za-z0-9\s&]{2,30})', subject, re.I)
        if m_subj_at:
            cand = clean_position(m_subj_at.group(1), company)
            if cand and TECH_TITLE_RE.search(cand):
                position = cand

    # 5. В теме письма: на польском "na stanowisko <Position>"
    if not position:
        m_pl_subj = re.search(r'na\s+stanowisko\s+([A-Za-z0-9\s\-_/,\(\)&|]+?)(?:\s*[-–—|]|\.|\n|$)', subject, re.I)
        if m_pl_subj:
            position = clean_position(m_pl_subj.group(1), company)

    # 6. В теме письма: на русском "на вакансию <Position>"
    if not position:
        m_ru_subj = re.search(r'на\s+вакансию\s+([A-Za-z0-9\s\-_/,\(\)&|«»\"“]+?)(?:\s*[-–—|]|\.|\n|$)', subject, re.I)
        if m_ru_subj:
            position = clean_position(m_ru_subj.group(1), company)

    # 7. В тексте письма (явные HR конструкции)
    if not position:
        patterns_body = [
            r'(?:role|position)\s+of\s+([A-Za-z0-9\s\-_/,\(\)&|]+?)(?:\s+(?:at|with|here\s+at)\b|\.|\n|$)',
            r'(?:applying|applied|application)\s+for\s+(?:the\s+)?([A-Za-z0-9\s\-_/,\(\)&|]+?)(?:\s+(?:role|position|job)\b|\s+(?:at|with|here\s+at)\b|\.|\n|$)',
            r'(?:interest in the|for the)\s+([A-Za-z0-9\s\-_/,\(\)&|]+?)\s+(?:role|position)\b',
            r'na\s+stanowisko\s+([A-Za-z0-9\s\-_/,\(\)&|]+?)(?:\s+(?:w|we)\b|\.|\n|$)',
            r'на\s+вакансию\s+([A-Za-z0-9\s\-_/,\(\)&|«»\"“]+?)(?:\s+(?:в|для)\b|\.|\n|$)'
        ]
        for pat in patterns_body:
            m_b = re.search(pat, body_one_line, re.I)
            if m_b:
                cand = clean_position(m_b.group(1), company)
                if cand:
                    position = cand
                    break

    # 8. Fallback: поиск стандартного технического тайтла в теме письма
    if not position:
        m_title = TECH_TITLE_RE.search(subject)
        if m_title:
            for part in re.split(r'[-–—|]', subject):
                if m_title.group(0).lower() in part.lower():
                    cand = clean_position(part, company)
                    if cand:
                        position = cand
                        break
            if not position:
                position = clean_position(m_title.group(0), company)

    company = re.sub(r'["«»]', '', company).strip()
    position = re.sub(r'["«»]', '', position).strip()

    return company or "Не определена", position or "Не указана"

def detect_email_status(subject: str, body: str):
    """
    Классифицирует статус письма: Оффер, Отказ, Приглашение, На рассмотрении.
    """
    # Анализируем только начало письма: дальше часто идёт копия резюме или описание вакансии
    full_text = f"{subject}\n{body[:1500]}".lower()

    # 1. Оффер (наивысший позитивный статус)
    for pattern in OFFER_PATTERNS:
        if re.search(pattern, full_text, re.IGNORECASE):
            return "Оффер (Offer)", pattern

    # 2. Отказ
    for pattern in REJECTION_PATTERNS:
        if re.search(pattern, full_text, re.IGNORECASE):
            return "Отказ (Rejected)", pattern

    # 3. Приглашение / Интервью
    if not re.search(r'job offer:\s+principal|prepare for interviews', full_text):
        for pattern in INTERVIEW_PATTERNS:
            if re.search(pattern, full_text, re.IGNORECASE):
                return "Приглашение (Interview)", pattern

    # 4. Подтверждение отклика
    for pattern in CONFIRMATION_PATTERNS:
        if re.search(pattern, full_text, re.IGNORECASE):
            return "На рассмотрении (In Review)", pattern

    return "На рассмотрении (In Review)", ""

def has_application_signal(subject: str, body: str) -> bool:
    """
    Сильный сигнал: в письме есть явный маркер отклика/отказа/интервью/оффера
    И контекст вакансии. Одних слов 'job'/'application' недостаточно.
    """
    text = f"{subject}\n{body[:1500]}"
    if not JOB_CONTEXT_RE.search(text):
        return False
    all_patterns = OFFER_PATTERNS + REJECTION_PATTERNS + INTERVIEW_PATTERNS + CONFIRMATION_PATTERNS
    return any(re.search(p, text, re.IGNORECASE) for p in all_patterns)

def parse_date(date_str: str) -> datetime:
    """Парсит заголовок даты письма."""
    if not date_str:
        return datetime.min
    try:
        dt = email.utils.parsedate_to_datetime(date_str)
        if dt.tzinfo:
            dt = dt.astimezone().replace(tzinfo=None)
        return dt
    except Exception:
        return datetime.min

def list_target_folders(mail):
    """
    Возвращает папки для сканирования: 'Вся почта' (флаг \\All), если она доступна по IMAP,
    иначе все выбираемые папки, кроме Корзины/Спама/Черновиков.
    """
    all_mail, folders = None, []
    status, lines = mail.list()
    for item in (lines or []):
        if not item:
            continue
        line = item.decode('latin1', errors='replace') if isinstance(item, bytes) else str(item)
        m = re.search(r'\((?P<flags>[^)]*)\)\s+"(?P<delim>[^"]+)"\s+(?P<name>.*)$', line)
        if not m:
            continue
        flags, name = m.group('flags'), m.group('name').strip()
        if r'\All' in flags:
            all_mail = name
        if any(f in flags for f in (r'\Noselect', r'\Trash', r'\Junk', r'\Drafts')):
            continue
        if name.strip('"').split('/')[-1].lower() in ('junk', 'trash', 'spam', 'deleted messages', 'drafts'):
            continue
        folders.append(name)
    return ([all_mail], True) if all_mail else (folders, False)

def get_imap_host(user_email: str, custom_host: str = "") -> str:
    """Определяет IMAP-сервер по домену почты или возвращает указанный пользователем."""
    if custom_host:
        return custom_host
    domain = user_email.split("@")[-1].lower() if "@" in user_email else ""
    if "yandex" in domain:
        return "imap.yandex.ru"
    if any(m in domain for m in ["mail.ru", "bk.ru", "inbox.ru", "list.ru"]):
        return "imap.mail.ru"
    if any(o in domain for o in ["outlook.com", "hotmail.com", "live.com"]):
        return "outlook.office365.com"
    if "yahoo" in domain:
        return "imap.mail.yahoo.com"
    return "imap.gmail.com"

def fetch_emails_from_account(user_email: str, app_password: str, since_date: str,
                              own_addresses: set, seen_ids: set, imap_server: str = ""):
    """Подключается к почтовому ящику по IMAP и выгружает письма начиная с даты."""
    host = get_imap_host(user_email, imap_server)
    print(f"\n[*] Подключение к {user_email} ({host})...")
    emails_data = []

    try:
        mail = imaplib.IMAP4_SSL(host)
        mail.login(user_email, app_password.replace(" ", ""))
    except Exception as e:
        print(f"[!] Не удалось подключиться к {user_email} через {host}: {e}")
        return []

    mail._encoding = "utf-8"
    folders, has_all_mail = list_target_folders(mail)
    if not has_all_mail:
        print("    ⚠ Папка 'Вся почта' скрыта от IMAP — сканирую все папки по отдельности.")
        print("      (Можно включить: Gmail → Настройки → Ярлыки → 'Вся почта' → 'Показывать в IMAP')")

    for folder in folders:
        try:
            status, _ = mail.select(folder, readonly=True)
            if status != 'OK':
                continue
            status, messages = mail.search(None, 'SINCE', since_date)
        except Exception:
            continue
        if status != 'OK' or not messages or not messages[0]:
            continue

        msg_ids = messages[0].split()

        # Быстрая загрузка заголовков пачками для мгновенного отсечения спама и рассылок
        candidate_ids = []
        chunking_ok = False
        chunk_size = 50
        for i in range(0, len(msg_ids), chunk_size):
            chunk = b','.join(msg_ids[i:i+chunk_size])
            try:
                res, data = mail.fetch(chunk, '(BODY.PEEK[HEADER.FIELDS (SUBJECT FROM TO DATE MESSAGE-ID IN-REPLY-TO REFERENCES)])')
                if res != 'OK' or not data:
                    continue
                chunking_ok = True
                for item in data:
                    if isinstance(item, tuple):
                        m_num = re.match(rb'^(\d+)\s+', item[0])
                        if not m_num:
                            continue
                        mid = m_num.group(1)
                        h_msg = email.message_from_bytes(item[1])
                        subj = decode_mime_words(h_msg.get("Subject", ""))
                        frm = decode_mime_words(h_msg.get("From", ""))
                        if is_noise_or_marketing(subj, frm, ""):
                            continue
                        candidate_ids.append(mid)
            except Exception:
                pass

        target_ids = candidate_ids if chunking_ok else msg_ids
        print(f"    {folder}: {len(msg_ids)} писем (к загрузке: {len(target_ids)})...", flush=True)
        for msg_id in target_ids:
            try:
                res, data = mail.fetch(msg_id, "(RFC822)")
                if res != 'OK' or not data or not data[0]:
                    continue
                msg = email.message_from_bytes(data[0][1])
                record = build_email_record(msg, user_email, own_addresses, seen_ids)
                if record:
                    emails_data.append(record)
            except Exception:
                continue

    try:
        mail.logout()
    except Exception:
        pass
    print(f"    Кандидатов в отчёт: {len(emails_data)}")
    return emails_data

def build_email_record(msg, account: str, own_addresses: set, seen_ids: set):
    """Разбирает письмо. Возвращает None для шума и дубликатов."""
    message_id = (msg.get("Message-ID") or "").strip()
    if message_id:
        if message_id in seen_ids:
            return None
        seen_ids.add(message_id)

    subject = decode_mime_words(msg.get("Subject", ""))
    from_hdr = decode_mime_words(msg.get("From", ""))
    to_hdr = decode_mime_words(msg.get("To", ""))
    body = get_body_text(msg)

    if is_noise_or_marketing(subject, from_hdr, body):
        return None

    is_mine = extract_email_address(from_hdr) in own_addresses
    if is_mine:
        # Своё письмо: компания определяется по получателю; статус не классифицируем
        company, position = parse_metadata_from_email(subject, to_hdr, body)
        status_text = "На рассмотрении (In Review)"
        strong = bool(JOB_CONTEXT_RE.search(subject))
    else:
        company, position = parse_metadata_from_email(subject, from_hdr, body)
        status_text, _ = detect_email_status(subject, body)
        strong = has_application_signal(subject, body)

    msg_dt = parse_date(msg.get("Date", ""))
    refs = f"{msg.get('In-Reply-To', '')} {msg.get('References', '')}"
    return {
        "account": account,
        "message_id_header": message_id,
        "refs": set(re.findall(r'<[^>]+>', refs)),
        "date": msg_dt,
        "date_str": msg_dt.strftime("%Y-%m-%d %H:%M") if msg_dt != datetime.min else "",
        "from": from_hdr,
        "subject": subject,
        "clean_subj": clean_subject_for_grouping(subject),
        "company": company,
        "norm_company": normalize_company_name(company),
        "position": position,
        "status": status_text,
        "strong": strong,
        "snippet": body.replace("\n", " ")[:250].strip(),
    }

def select_relevant(all_emails):
    """
    Оставляет письма с сильным сигналом и всю переписку вокруг них
    (ответы по References/In-Reply-To или по совпадающей теме Re:/Odp:).
    """
    relevant = [e for e in all_emails if e["strong"]]
    rest = [e for e in all_emails if not e["strong"]]
    changed = True
    while changed:
        changed = False
        ids = {e["message_id_header"] for e in relevant}
        subjects = {e["clean_subj"].lower() for e in relevant if len(e["clean_subj"]) > 5}
        still_rest = []
        for e in rest:
            if (e["refs"] & ids) or e["clean_subj"].lower() in subjects:
                relevant.append(e)
                changed = True
            else:
                still_rest.append(e)
        rest = still_rest
    return relevant

def aggregate_job_applications(all_emails):
    """
    Интеллектуальная агрегация откликов:
    Связывает письма одного работодателя между обоими ящиками, вычисляет
    дни ожидания и сохраняет историю взаимодействия.
    """
    company_groups = defaultdict(list)
    orphans = []

    for item in all_emails:
        norm_c = item["norm_company"]
        if norm_c and len(norm_c) >= 2:
            company_groups[norm_c].append(item)
        else:
            orphans.append(item)

    # Привязка писем без определенной компании к существующим группам
    remaining_orphans = []
    for orphan in orphans:
        matched = False
        orphan_text = f"{orphan['clean_subj']} {orphan['from']} {orphan['snippet']}".lower()

        for norm_c in list(company_groups.keys()):
            if norm_c in orphan_text:
                company_groups[norm_c].append(orphan)
                matched = True
                break

        if not matched:
            for norm_c, grp_emails in company_groups.items():
                if any(e["clean_subj"].lower() == orphan["clean_subj"].lower() for e in grp_emails):
                    company_groups[norm_c].append(orphan)
                    matched = True
                    break

        if not matched:
            remaining_orphans.append(orphan)

    subject_groups = defaultdict(list)
    for orphan in remaining_orphans:
        subj_key = orphan["clean_subj"].lower()
        subject_groups[subj_key].append(orphan)

    all_clusters = list(company_groups.values()) + list(subject_groups.values())
    aggregated_records = []
    now = datetime.now()

    for thread_emails in all_clusters:
        thread_emails.sort(key=lambda x: x["date"])

        first_email = thread_emails[0]
        last_email = thread_emails[-1]

        valid_companies = [e["company"] for e in thread_emails if e["company"] != "Не определена"]
        best_company = max(valid_companies, key=len) if valid_companies else first_email["company"]

        generic_pos_words = ["не указана", "quick check-in", "feedback", "update", "application", "stanowisko", "your application"]
        pos_candidates = [
            e["position"] for e in thread_emails 
            if e["position"] and not any(g in e["position"].lower() for g in generic_pos_words)
        ]
        best_position = pos_candidates[0] if pos_candidates else first_email["position"]

        # Иерархия статусов: Оффер > Отказ > Интервью > На рассмотрении
        rejection_emails = [e for e in thread_emails if "Отказ" in e["status"]]
        interview_emails = [e for e in thread_emails if "Приглашение" in e["status"]]
        offer_emails = [e for e in thread_emails if "Оффер" in e["status"]]

        if offer_emails:
            decisive = offer_emails[-1]
            final_status = "🎉 Оффер (Offer)"
        elif rejection_emails:
            decisive = rejection_emails[-1]
            final_status = "Отказ (Rejected)"
        elif interview_emails:
            decisive = interview_emails[-1]
            final_status = "Приглашение (Interview)"
        else:
            decisive = last_email
            final_status = "На рассмотрении (In Review)"

        status_date = decisive["date_str"]
        status_details = decisive["snippet"]
        last_subj = decisive["subject"]
        last_sender = decisive["from"]

        # Расчет дней с последнего обновления
        last_dt = decisive["date"]
        days_ago = (now - last_dt).days if last_dt != datetime.min else 0
        days_ago_str = f"{days_ago} дн." if days_ago > 0 else "сегодня"

        accounts_involved = sorted(list(set(e["account"] for e in thread_emails)))
        accounts_str = ", ".join(accounts_involved)

        # Ручная корректировка статуса из overrides.json (например, отказ сообщили по телефону)
        override = USER_STATUS_OVERRIDES.get(normalize_company_name(best_company))
        if override:
            final_status = override

        aggregated_records.append({
            "account": accounts_str,
            "company": best_company,
            "position": best_position,
            "first_contact_date": first_email["date_str"],
            "status": final_status,
            "last_response_date": status_date,
            "days_ago": days_ago_str,
            "days_num": days_ago,
            "last_subject": last_subj,
            "sender": last_sender,
            "total_emails_in_thread": len(thread_emails),
            "snippet": status_details
        })

    # Сортировка: сначала свежие отклики
    aggregated_records.sort(key=lambda x: x["first_contact_date"], reverse=True)
    return aggregated_records

def export_to_csv(records, filename="job_applications_report.csv"):
    """Экспорт в CSV с UTF-8 BOM для Excel / Numbers / Sheets."""
    fieldnames = [
        "Почтовый ящик",
        "Компания",
        "Позиция / Вакансия",
        "Дата подачи (отклика)",
        "Статус",
        "Дата ответа / обновления",
        "Дней без ответа",
        "Тема последнего письма",
        "Отправитель",
        "Писем в ветке",
        "Фрагмент текста"
    ]

    with open(filename, mode="w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(fieldnames)
        for r in records:
            writer.writerow([
                r["account"],
                r["company"],
                r["position"],
                r["first_contact_date"],
                r["status"],
                r["last_response_date"],
                r["days_ago"],
                r["last_subject"],
                r["sender"],
                r["total_emails_in_thread"],
                r["snippet"]
            ])
    print(f"[+] CSV отчет сохранен: {os.path.abspath(filename)}")

def export_to_html(records, filename="job_applications_report.html"):
    """Генерирует адаптивный веб-дашборд с фильтрацией и аналитикой."""
    total_apps = len(records)
    total_rejected = sum(1 for r in records if "Отказ" in r["status"])
    total_interviews = sum(1 for r in records if "Приглашение" in r["status"])
    total_offers = sum(1 for r in records if "Оффер" in r["status"])
    total_in_review = sum(1 for r in records if "рассмотрении" in r["status"])

    html = f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Job Search Applications Tracker</title>
<style>
  :root {{
    --bg-primary: #0f172a;
    --bg-secondary: #1e293b;
    --bg-card: #1e293b;
    --text-primary: #f8fafc;
    --text-muted: #94a3b8;
    --border: #334155;
    --accent: #38bdf8;
    --badge-reject-bg: #450a0a;
    --badge-reject-text: #f87171;
    --badge-interview-bg: #064e3b;
    --badge-interview-text: #34d399;
    --badge-review-bg: #451a03;
    --badge-review-text: #fbbf24;
    --badge-offer-bg: #14532d;
    --badge-offer-text: #86efac;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    background-color: var(--bg-primary);
    color: var(--text-primary);
    padding: 24px;
    line-height: 1.5;
  }}
  .container {{ max-width: 1400px; margin: 0 auto; }}
  header {{ margin-bottom: 24px; display: flex; justify-content: space-between; align-items: flex-end; flex-wrap: wrap; gap: 16px; }}
  h1 {{ font-size: 26px; font-weight: 700; color: #fff; display: flex; align-items: center; gap: 10px; }}
  .subtitle {{ color: var(--text-muted); font-size: 14px; margin-top: 4px; }}
  
  .stats-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 16px;
    margin-bottom: 24px;
  }}
  .stat-card {{
    background: var(--bg-card);
    border: 1px solid var(--border);
    padding: 16px 20px;
    border-radius: 12px;
    box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2);
  }}
  .stat-title {{ font-size: 13px; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.05em; }}
  .stat-value {{ font-size: 28px; font-weight: 700; margin-top: 4px; }}
  .val-total {{ color: var(--accent); }}
  .val-reject {{ color: var(--badge-reject-text); }}
  .val-interview {{ color: var(--badge-interview-text); }}
  .val-review {{ color: var(--badge-review-text); }}
  .val-offer {{ color: var(--badge-offer-text); }}

  .controls {{
    display: flex;
    gap: 12px;
    margin-bottom: 16px;
    flex-wrap: wrap;
    align-items: center;
  }}
  .search-input {{
    flex: 1;
    min-width: 260px;
    padding: 10px 14px;
    border-radius: 8px;
    border: 1px solid var(--border);
    background: var(--bg-secondary);
    color: #fff;
    font-size: 14px;
    outline: none;
  }}
  .search-input:focus {{ border-color: var(--accent); }}
  .filter-btn {{
    padding: 8px 16px;
    border-radius: 8px;
    border: 1px solid var(--border);
    background: var(--bg-secondary);
    color: var(--text-muted);
    font-size: 13px;
    cursor: pointer;
    font-weight: 500;
    transition: all 0.2s;
  }}
  .filter-btn:hover, .filter-btn.active {{
    background: var(--accent);
    color: #0f172a;
    border-color: var(--accent);
  }}

  .table-wrapper {{
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: 12px;
    overflow-x: auto;
    box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2);
  }}
  table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 14px;
    text-align: left;
  }}
  th {{
    background: #1e293b;
    padding: 12px 16px;
    font-size: 12px;
    color: var(--text-muted);
    text-transform: uppercase;
    letter-spacing: 0.05em;
    border-bottom: 1px solid var(--border);
    white-space: nowrap;
  }}
  td {{
    padding: 14px 16px;
    border-bottom: 1px solid var(--border);
    vertical-align: top;
  }}
  tr:last-child td {{ border-bottom: none; }}
  tr:hover td {{ background: rgba(255, 255, 255, 0.02); }}

  .badge {{
    display: inline-block;
    padding: 4px 10px;
    border-radius: 9999px;
    font-size: 12px;
    font-weight: 600;
    white-space: nowrap;
  }}
  .badge-reject {{ background: var(--badge-reject-bg); color: var(--badge-reject-text); border: 1px solid #7f1d1d; }}
  .badge-interview {{ background: var(--badge-interview-bg); color: var(--badge-interview-text); border: 1px solid #065f46; }}
  .badge-review {{ background: var(--badge-review-bg); color: var(--badge-review-text); border: 1px solid #78350f; }}
  .badge-offer {{ background: var(--badge-offer-bg); color: var(--badge-offer-text); border: 1px solid #16a34a; }}

  .company-cell {{ font-weight: 600; color: #fff; font-size: 15px; }}
  .position-cell {{ color: #cbd5e1; font-weight: 500; }}
  .date-cell {{ color: var(--text-muted); font-size: 13px; white-space: nowrap; }}
  .days-pill {{
    display: inline-block;
    background: rgba(255, 255, 255, 0.06);
    padding: 2px 6px;
    border-radius: 4px;
    font-size: 11px;
    color: #cbd5e1;
    margin-top: 2px;
  }}
  .account-pill {{
    display: inline-block;
    background: #0f172a;
    border: 1px solid var(--border);
    padding: 2px 8px;
    border-radius: 6px;
    font-size: 11px;
    color: #94a3b8;
    margin-top: 4px;
  }}
  .thread-count {{
    display: inline-block;
    background: #334155;
    padding: 1px 6px;
    border-radius: 4px;
    font-size: 10px;
    color: #cbd5e1;
    margin-left: 6px;
  }}
  .snippet-text {{
    font-size: 12px;
    color: var(--text-muted);
    margin-top: 6px;
    max-width: 480px;
    display: -webkit-box;
    -webkit-line-clamp: 2;
    -webkit-box-orient: vertical;
    overflow: hidden;
  }}
</style>
</head>
<body>
<div class="container">
  <header>
    <div>
      <h1><span>💼</span> Трекер откликов на вакансии</h1>
      <p class="subtitle">Отслеживание вакансий | Всего работодателей в пайплайне: {total_apps}</p>
    </div>
    <div>
      <span style="font-size: 12px; color: var(--text-muted);">Обновлено: {datetime.now().strftime('%d.%m.%Y %H:%M')}</span>
    </div>
  </header>

  <div class="stats-grid">
    <div class="stat-card">
      <div class="stat-title">Всего компаний</div>
      <div class="stat-value val-total">{total_apps}</div>
    </div>
    <div class="stat-card">
      <div class="stat-title">Отказы</div>
      <div class="stat-value val-reject">{total_rejected}</div>
    </div>
    <div class="stat-card">
      <div class="stat-title">Интервью / Тестирование</div>
      <div class="stat-value val-interview">{total_interviews}</div>
    </div>
    <div class="stat-card">
      <div class="stat-title">На рассмотрении</div>
      <div class="stat-value val-review">{total_in_review}</div>
    </div>
  </div>

  <div class="controls">
    <input type="text" id="searchInput" class="search-input" placeholder="🔍 Поиск по компании, позиции или теме..." onkeyup="filterRows()">
    <button class="filter-btn active" onclick="setFilter('all', this)">Все ({total_apps})</button>
    <button class="filter-btn" onclick="setFilter('Отказ', this)">Отказы ({total_rejected})</button>
    <button class="filter-btn" onclick="setFilter('Приглашение', this)">Приглашения ({total_interviews})</button>
    <button class="filter-btn" onclick="setFilter('рассмотрении', this)">На рассмотрении ({total_in_review})</button>
  </div>

  <div class="table-wrapper">
    <table id="appsTable">
      <thead>
        <tr>
          <th>Компания & Ящик</th>
          <th>Позиция</th>
          <th>Дата отклика</th>
          <th>Статус</th>
          <th>Обновление</th>
          <th>Последняя тема и фрагмент</th>
        </tr>
      </thead>
      <tbody>
"""

    for r in records:
        status_badge_class = "badge-review"
        if "Оффер" in r["status"]:
            status_badge_class = "badge-offer"
        elif "Отказ" in r["status"]:
            status_badge_class = "badge-reject"
        elif "Приглашение" in r["status"]:
            status_badge_class = "badge-interview"

        thread_badge = f'<span class="thread-count">{r["total_emails_in_thread"]} писем</span>' if r["total_emails_in_thread"] > 1 else ''

        html += f"""
        <tr data-status="{escape(r['status'])}">
          <td>
            <div class="company-cell">{escape(r['company'])} {thread_badge}</div>
            <div class="account-pill">{escape(r['account'])}</div>
          </td>
          <td class="position-cell">{escape(r['position'])}</td>
          <td class="date-cell">{escape(r['first_contact_date'])}</td>
          <td><span class="badge {status_badge_class}">{escape(r['status'])}</span></td>
          <td class="date-cell">
            <div>{escape(r['last_response_date'])}</div>
            <div class="days-pill">⏳ {escape(r['days_ago'])}</div>
          </td>
          <td>
            <div style="font-weight: 500; font-size: 13px;">{escape(r['last_subject'])}</div>
            <div class="snippet-text">{escape(r['snippet'])}</div>
          </td>
        </tr>
"""

    html += """
      </tbody>
    </table>
  </div>
</div>

<script>
  let currentFilter = 'all';

  function setFilter(status, btn) {
    currentFilter = status;
    document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    filterRows();
  }

  function filterRows() {
    const q = document.getElementById('searchInput').value.toLowerCase();
    const rows = document.querySelectorAll('#appsTable tbody tr');

    rows.forEach(row => {
      const text = row.innerText.toLowerCase();
      const rowStatus = row.getAttribute('data-status') || '';
      
      const matchesSearch = text.includes(q);
      const matchesFilter = currentFilter === 'all' || rowStatus.includes(currentFilter);

      if (matchesSearch && matchesFilter) {
        row.style.display = '';
      } else {
        row.style.display = 'none';
      }
    });
  }
</script>
</body>
</html>
"""
    with open(filename, mode="w", encoding="utf-8") as f:
        f.write(html)
    print(f"[+] HTML интерактивный отчет сохранен: {os.path.abspath(filename)}")

def load_or_request_accounts(config_file="accounts.json"):
    """Загружает сохраненные учетные записи или запрашивает новые."""
    if os.path.exists(config_file):
        try:
            with open(config_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list) and len(data) > 0:
                    print(f"[*] Загружены настройки из {config_file} ({len(data)} аккаунт(ов)).")
                    return data
        except Exception as e:
            print(f"[!] Не удалось прочитать {config_file}: {e}")

    print("\n" + "="*70)
    print("НАСТРОЙКА ПОДКЛЮЧЕНИЯ К ПОЧТОВЫМ ЯЩИКАМ GMAIL")
    print("="*70)
    print("Для подключения требуется 'Пароль приложения' (App Password).")
    print("Получить: https://myaccount.google.com/apppasswords")
    print("="*70 + "\n")

    accounts = []
    i = 1
    while True:
        prompt_txt = f"Введите email адрес #{i} (или Enter для завершения): " if i > 1 else "Введите email адрес (например, yourname@gmail.com): "
        email_addr = input(prompt_txt).strip()
        if not email_addr:
            break
        password = getpass.getpass(f"Введите пароль приложения для {email_addr}: ").strip()
        if password:
            accounts.append({"email": email_addr, "app_password": password})
        i += 1

    save = input("\nСохранить учетные записи в accounts.json? (y/n) [y]: ").strip().lower()
    if save in ["", "y", "yes", "да"]:
        with open(config_file, "w", encoding="utf-8") as f:
            json.dump(accounts, f, indent=2, ensure_ascii=False)
        print(f"[+] Сохранено в {config_file}.")

    return accounts

IGNORE_TEMPLATE = """# Исключения: по одной строке. Если строка встречается в отправителе или теме письма
# (без учёта регистра), письмо не попадёт в отчёт. Примеры:
# @simplify.jobs
# newsletter@
"""

def load_user_rules(base_dir):
    """Загружает ignore.txt и overrides.json (создаёт шаблон ignore.txt при отсутствии)."""
    ignore_path = os.path.join(base_dir, "ignore.txt")
    if not os.path.exists(ignore_path):
        with open(ignore_path, "w", encoding="utf-8") as f:
            f.write(IGNORE_TEMPLATE)
    with open(ignore_path, encoding="utf-8") as f:
        USER_IGNORE_PATTERNS.extend(
            line.strip().lower() for line in f if line.strip() and not line.startswith("#")
        )

    overrides_path = os.path.join(base_dir, "overrides.json")
    if os.path.exists(overrides_path):
        with open(overrides_path, encoding="utf-8") as f:
            for company, status in json.load(f).items():
                USER_STATUS_OVERRIDES[normalize_company_name(company)] = status

def main():
    parser = argparse.ArgumentParser(description="Job Search Email Tracker & Parser")
    parser.add_argument("--since", default=DEFAULT_SINCE_DATE, help=f"Дата начала поиска в формате DD-Mon-YYYY (по умолчанию: {DEFAULT_SINCE_DATE})")
    parser.add_argument("--csv", default="job_applications_report.csv", help="Путь для сохранения CSV файла")
    parser.add_argument("--html", default="job_applications_report.html", help="Путь для сохранения HTML отчета")
    args = parser.parse_args()

    print("="*70)
    print("  JOB SEARCH EMAIL TRACKER & PIPELINE MANAGER")
    print("="*70)

    config_path = os.path.join(os.path.dirname(__file__), "accounts.json")
    accounts = load_or_request_accounts(config_path)
    load_user_rules(os.path.dirname(os.path.abspath(__file__)))

    if not accounts:
        print("[!] Нет настроенных аккаунтов. Выход.")
        sys.exit(1)

    all_emails = []
    own_addresses = {acc["email"].lower() for acc in accounts if acc.get("email")}
    seen_ids = set()
    for acc in accounts:
        user_email = acc.get("email")
        app_password = acc.get("app_password")
        imap_server = acc.get("imap_server", "")
        if not user_email or not app_password:
            continue
        emails = fetch_emails_from_account(user_email, app_password, args.since, own_addresses, seen_ids, imap_server)
        all_emails.extend(emails)

    all_emails = select_relevant(all_emails)
    print(f"\n{'='*70}")
    print(f"Всего отобрано релевантных писем по вакансиям: {len(all_emails)}")
    print(f"Группировка и связывание откликов и ответов...")

    records = aggregate_job_applications(all_emails)
    print(f"Сформировано уникальных работодателей в пайплайне: {len(records)}")

    csv_file = os.path.join(os.path.dirname(__file__), args.csv)
    html_file = os.path.join(os.path.dirname(__file__), args.html)

    export_to_csv(records, csv_file)
    export_to_html(records, html_file)

    print("\n[✔] Готово! Результаты актуализированы:")
    print(f"    - Таблица: {csv_file}")
    print(f"    - Интерактивный дашборд: {html_file}")

if __name__ == "__main__":
    main()
