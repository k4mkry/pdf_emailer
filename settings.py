import tkinter as tk
from tkinter import ttk
from datetime import datetime


class Style(ttk.Style):
    def __init__(self, root, **kw):
        ttk.Style.__init__(self, master=root, **kw)
        self.configure(
            'warning.TLabel',
            foreground='red4',
            font=('Calibri', 12, 'bold'),
            anchor='center'
        )
        self.configure(
            'info.TLabel',
            foreground='dark green',
            font=('Calibri', 12, 'bold'),
            anchor='center'
        )


class Settings:

    DIRECTORY = r'C:\Users\kkrysa.HMT\Documents\Dev\python\tkinter\AKS_invoice_emailer\WDT_path'
    # DIRECTORY = r'C:\Users\kkrysa.HMT\Desktop\Dev\Python\AKS_invoice_emailer\WDT_path'
    date_now = datetime.now().strftime('%Y_%m_%d') + '-'
    body = '''
    <p>
    <span style="font-size:8px">[de] </span>
    Sehr geehrte Damen und Herren,
    </p>
    <p>
    Im Anhang finden Sie unsere elektronische Rechnung. Nachricht wurde automatisch generiert - bitte antworten Sie nicht darauf.
    </p>
    <br>

    <p>
    <span style="font-size:8px">[en] </span>
    Dear Sirs,
    </p>
    <p>
    Please find attached our invoice. Message is generated automatically  - please do not reply to it.
    </p>
    <br>

    <p>
    <span style="font-size:8px">[pl] </span>
    Dzień Dobry,
    </p>
    <p>
    W załączeniu przesyłamy naszą fakturę w wersji elektronicznej. Wiadomość została wygenerowana automatycznie – proszę na nią nie odpowiadać.
    </p>
    <br>

    <p><small><i>
    HMT Heldener Metalltechnik Polska Sp. z o.o. & Co. Sp. K.<br>
    ul. Polna 17A  , Komorniki<br>
    55-300 Środa Śląska<br>
    Tel.:+48  71 74 72 961<br>
    Fax: +48 71 74 72 901<br>
    <a href="www.hmt-automotive.com">www.hmt-automotive.com</a>
    </i></small></p>
    '''
