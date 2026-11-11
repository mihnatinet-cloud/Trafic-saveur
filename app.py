# -*- coding: utf-8 -*-
"""
Trafik Saveurs - Système de Gestion Commerciale & de Stock
نظام إدارة المبيعات والمخزون والتوزيع
Version: 3.0 - مع شاشة تحميل عصرية (Splash Screen)
"""
import os, sys, sqlite3, csv, shutil, webbrowser, html, ctypes, threading, time
from pathlib import Path
from datetime import datetime, date, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# تعيين ID التطبيق لويندوز لعرض الأيقونة الصحيحة
if sys.platform == "win32":
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("trafiksaveurs.commercial.v3")
    except Exception:
        pass

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import winsound
except ImportError:
    winsound = None

# ============================================================
# إعدادات المسارات والألوان
# ============================================================
APP_NAME = "Trafik Saveurs"
APP_SUBTITLE_FR = "Gestion de Distribution de Biscuiterie & Confiserie en Gros"
APP_SUBTITLE_AR = "نظام إدارة توزيع البسكويت والحلويات بالجملة"
ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
BACKUP_DIR = ROOT / "sauvegardes"
DATA_DIR.mkdir(exist_ok=True)
BACKUP_DIR.mkdir(exist_ok=True)
DB_PATH = DATA_DIR / "trafik_saveurs.db"
ICON_PATH = ROOT / "icon.ico"
LOGO_PATH = ROOT / "images" / "1794164544.png"

if not LOGO_PATH.exists():
    alt_logo = ROOT / "Affiche_Trafik_Saveurs_reference.png"
    if alt_logo.exists():
        LOGO_PATH = alt_logo
    else:
        # Logo intégré (généré localement) — application 100% hors ligne
        LOGO_PATH = ROOT / "icon.png"

# لوحة الألوان الاحترافية
BG = "#F6F3EC"
PANEL = "#FFFFFF"
HEADER_BG = "#29130D"
BROWN = "#3B1B12"
GOLD = "#E09C00"
GOLD_HOVER = "#C78800"
MUTED = "#6E6259"
BORDER = "#E0D7CB"
GREEN = "#217A4B"
RED = "#B83226"
ORANGE = "#D96E14"
BLUE = "#1B6396"

def beep_ok():
    if winsound:
        try: winsound.MessageBeep(winsound.MB_OK)
        except Exception: pass

def beep_err():
    if winsound:
        try: winsound.MessageBeep(winsound.MB_ICONHAND)
        except Exception: pass

_GLOBAL_ICON_PHOTO = None

def apply_window_icon(win):
    global _GLOBAL_ICON_PHOTO
    if ICON_PATH.exists():
        try:
            win.iconbitmap(str(ICON_PATH))
        except Exception:
            pass
        try:
            if PIL_AVAILABLE:
                if _GLOBAL_ICON_PHOTO is None:
                    _GLOBAL_ICON_PHOTO = ImageTk.PhotoImage(file=str(ICON_PATH))
                win.iconphoto(True, _GLOBAL_ICON_PHOTO)
        except Exception:
            pass

def center_window(win, width, height):
    apply_window_icon(win)
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    max_w = max(380, int(sw * 0.95))
    max_h = max(280, int(sh * 0.82))
    w = min(width, max_w)
    h = min(height, max_h)
    x = max(0, (sw - w) // 2)
    y = max(0, (sh - h) // 2)
    win.geometry(f"{w}x{h}+{x}+{y}")

# ============================================================
# القاموس ثنائي اللغة
# ============================================================
I18N = {
    "fr": {
        "app_title": "Trafik Saveurs | Gestion Commerciale & Stock",
        "subtitle": "Gestion de Distribution de Biscuiterie & Confiserie en Gros",
        "offline_badge": "Système Local • Sécurisé Hors Ligne",
        "curr": "DA",
        "dashboard": "Tableau de Bord",
        "sales": "Ventes",
        "products": "Produits & Stock",
        "customers": "Clients",
        "purchases": "Achats",
        "expenses": "Dépenses",
        "reports": "Rapports",
        "settings": "Paramètres",
        "backup": "Sauvegarde",
        "logout": "Déconnexion",
        "user": "Utilisateur",
        "role": "Rôle",
        "today_sales": "Ventes du Jour",
        "today_paid": "Encaissements du Jour",
        "products_count": "Total Produits",
        "critical_stock": "Stock Critique",
        "expiring_soon": "Péremption Proche",
        "customer_debts": "Créances Clients",
        "recent_sales": "Dernières Ventes",
        "alert_section": "Alertes Stock & Expiration",
        "invoice_num": "N° Facture",
        "date": "Date",
        "customer": "Client",
        "total": "Total",
        "paid": "Payé",
        "due": "Reste",
        "add_product": "+ Ajouter Produit",
        "edit_product": "✏️ Modifier Produit",
        "delete_product": "🗑️ Supprimer Produit",
        "clear_products_table": "⚠️ Vider le Tableau",
        "print_products": "🖨️ Imprimer Liste Produits",
        "search": "Rechercher",
        "designation": "Désignation",
        "category": "Catégorie",
        "barcode": "Code-barres",
        "charg": "Charg.",
        "purchase_price": "Prix d'Achat",
        "sale_price": "Prix de Vente",
        "stock": "Stock",
        "unit": "Unité",
        "alert_threshold": "Seuil d'Alerte",
        "expiry_date": "Date d'Expiration",
        "new_invoice": "+ Nouvelle Facture",
        "edit_sale": "✏️ Modifier Facture",
        "delete_sale": "🗑️ Supprimer Facture",
        "clear_sales_table": "⚠️ Vider le Tableau",
        "print_a4": "🖨️ Imprimer Facture A4",
        "print_ticket": "🧾 Imprimer Ticket 80mm",
        "print_sales_list": "🖨️ Imprimer Journal Ventes",
        "scan_barcode": "Scanner ou saisir Code-barres...",
        "quantity": "Quantité (Total)",
        "unit_price": "Prix Uni.",
        "total_produit": "Total Produit",
        "total_net": "TOTAL NET",
        "total_products": "Total Produits",
        "add_to_cart": "Ajouter au Panier",
        "delete_line": "Supprimer Ligne",
        "pay_full": "Tout Payer",
        "pay_zero": "Crédit Total",
        "payment_method": "Mode de Paiement",
        "notes": "Notes / Remarques",
        "save_sale": "Enregistrer la Vente et Déduire du Stock",
        "counter_customer": "Client Comptoir",
        "add_customer": "+ Ajouter Client",
        "edit_customer": "✏️ Modifier Client",
        "delete_customer": "🗑️ Supprimer Client",
        "pay_debt": "💵 Encaisser Paiement",
        "account_statement": "📋 Relevé de Compte",
        "print_customers_list": "🖨️ Imprimer Liste Clients",
        "phone": "Téléphone",
        "address": "Adresse",
        "balance_due": "Solde Dû",
        "new_purchase": "+ Enregistrer Achat",
        "import_excel": "📥 Importer Excel (Achats & Stock)",
        "download_template": "📄 Modèle Excel",
        "edit_purchase": "✏️ Modifier Achat",
        "delete_purchase": "🗑️ Supprimer Achat",
        "clear_purchases_table": "⚠️ Vider le Tableau",
        "print_purchases_list": "🖨️ Imprimer Journal Achats",
        "supplier": "Fournisseur",
        "unit_cost": "Coût Unitaire",
        "add_expense": "+ Ajouter Dépense",
        "edit_expense": "✏️ Modifier Dépense",
        "delete_expense": "🗑️ Supprimer Dépense",
        "print_expenses_list": "🖨️ Imprimer Journal Dépenses",
        "expense_title": "Libellé de la Dépense",
        "amount": "Montant",
        "from_date": "Du",
        "to_date": "Au",
        "view_report": "Afficher Rapport",
        "export_csv": "Exporter CSV",
        "export_excel": "📗 Exporter Excel",
        "import_customers_excel": "📥 Importer Clients (Excel)",
        "download_customers_template": "📄 Modèle Clients",
        "excel_missing": "Bibliothèque 'openpyxl' introuvable.\nInstallez-la une fois avec : pip install openpyxl",
        "total_sales": "Total Ventes",
        "total_purchases": "Total Achats",
        "cogs": "Coût des Marchandises Vendues",
        "gross_margin": "Marge Brute",
        "total_expenses": "Total Dépenses",
        "net_profit": "Bénéfice Net Estimé",
        "users_tab": "Gestion des Utilisateurs & Mots de Passe",
        "company_tab": "Informations de l'Établissement & Facturation",
        "add_user": "+ Ajouter Utilisateur",
        "edit_user": "✏️ Modifier Utilisateur",
        "change_pwd": "🔑 Changer Mot de Passe",
        "delete_user": "🗑️ Supprimer Utilisateur",
        "print_users_list": "🖨️ Imprimer Liste Utilisateurs",
        "username": "Nom d'utilisateur",
        "password": "Mot de passe",
        "confirm_pwd": "Confirmer Mot de passe",
        "admin": "Administrateur",
        "cashier": "Vendeur / Caissier",
        "save_settings": "Enregistrer les Paramètres",
        "company_name": "Nom de l'Entreprise / Magasin",
        "company_activity": "Activité Commerciale",
        "company_phone": "Téléphone Commercial",
        "company_address": "Adresse Complète",
        "rc": "N° Registre de Commerce (RC)",
        "nif": "N° Identification Fiscale (NIF)",
        "invoice_format": "Format d'Impression par Défaut",
        "format_a4": "Facture A4",
        "format_ticket": "Ticket de Caisse 80mm",
        "invoice_footer_msg": "Message de Bas de Page Facture",
        "confirm_delete": "Voulez-vous vraiment supprimer cet élément ?",
        "saved_success": "Opération effectuée avec succès.",
        "field_required": "Veuillez remplir tous les champs obligatoires.",
        "pwd_mismatch": "Les mots de passe ne correspondent pas.",
        "stock_insufficient": "Quantité demandée supérieure au stock disponible.",
        "product_not_found": "Produit introuvable pour ce code-barres.",
        "cannot_delete_self": "Vous ne pouvez pas supprimer votre propre compte connecté.",
        "cannot_delete_last_admin": "Impossible de supprimer le dernier administrateur.",
        "customer_debt_warning": "Ce client possède des dettes impayées. Voulez-vous vraiment le supprimer ?",
        "sale_delete_warning": "Cette facture sera supprimée et les quantités vendues seront restituées au stock. Continuer ?",
        "purchase_delete_warning": "Cet achat sera supprimé et les quantités seront déduites du stock. Continuer ?",
    },
    "ar": {
        "app_title": "ترافيك سافور | نظام إدارة المبيعات والمخزون",
        "subtitle": "إدارة توزيع البسكويت والحلويات بالجملة",
        "offline_badge": "نظام محلي آمن • يعمل بدون إنترنت",
        "curr": "د.ج",
        "dashboard": "لوحة التحكم",
        "sales": "المبيعات",
        "products": "المنتجات والمخزون",
        "customers": "الزبائن",
        "purchases": "المشتريات",
        "expenses": "المصاريف",
        "reports": "التقارير",
        "settings": "الإعدادات",
        "backup": "نسخ احتياطي",
        "logout": "تسجيل الخروج",
        "user": "المستخدم",
        "role": "الصلاحية",
        "today_sales": "مبيعات اليوم",
        "today_paid": "تحصيلات اليوم",
        "products_count": "إجمالي المنتجات",
        "critical_stock": "تنبيه المخزون",
        "expiring_soon": "قريب الصلاحية",
        "customer_debts": "ديون الزبائن",
        "recent_sales": "آخر الفواتير",
        "alert_section": "تنبيهات المخزون وانتهاء الصلاحية",
        "invoice_num": "رقم الفاتورة",
        "date": "التاريخ",
        "customer": "الزبون",
        "total": "المجموع",
        "paid": "المدفوع",
        "due": "المتبقي",
        "add_product": "+ إضافة منتج جديد",
        "edit_product": "✏️ تعديل المنتج",
        "delete_product": "🗑️ حذف المنتج",
        "clear_products_table": "⚠️ تفريغ جدول المنتجات",
        "print_products": "🖨️ طباعة قائمة المنتجات",
        "search": "بحث",
        "designation": "اسم المنتج (Désignation)",
        "category": "الفئة / الصنف",
        "barcode": "الباركود",
        "charg": "الحمولة (Charg)",
        "purchase_price": "سعر الشراء",
        "sale_price": "سعر البيع",
        "stock": "المخزون الحالي",
        "unit": "الوحدة",
        "alert_threshold": "حد التنبيه",
        "expiry_date": "تاريخ نهاية الصلاحية",
        "new_invoice": "+ فاتورة بيع جديدة",
        "edit_sale": "✏️ تعديل الفاتورة",
        "delete_sale": "🗑️ حذف الفاتورة",
        "clear_sales_table": "⚠️ تفريغ جدول المبيعات",
        "print_a4": "🖨️ طباعة فاتورة A4",
        "print_ticket": "🧾 طباعة وصل 80mm",
        "print_sales_list": "🖨️ طباعة سجل المبيعات",
        "scan_barcode": "امسح أو اكتب الباركود...",
        "quantity": "الكمية الإجمالية (Total)",
        "unit_price": "سعر الوحدة (Prix Uni)",
        "total_produit": "المجموع (Total Produit)",
        "total_net": "المجموع الصافي (TOTAL NET)",
        "total_products": "مجموع المنتجات",
        "add_to_cart": "إضافة إلى السلة",
        "delete_line": "حذف السطر",
        "pay_full": "دفع كامل",
        "pay_zero": "دين كامل",
        "payment_method": "طريقة الدفع",
        "notes": "ملاحظات",
        "save_sale": "تسجيل الفاتورة وخصم المخزون",
        "counter_customer": "زبون عادي",
        "add_customer": "+ إضافة زبون",
        "edit_customer": "✏️ تعديل الزبون",
        "delete_customer": "🗑️ حذف الزبون",
        "pay_debt": "💵 تسديد دين",
        "account_statement": "📋 كشف حساب زبون",
        "print_customers_list": "🖨️ طباعة قائمة الزبائن",
        "phone": "الهاتف",
        "address": "العنوان",
        "balance_due": "الرصيد المتبقي",
        "new_purchase": "+ تسجيل فاتورة شراء",
        "import_excel": "📥 استيراد إكسل (مشتريات ومخزون)",
        "download_template": "📄 نموذج إكسل جاهز",
        "edit_purchase": "✏️ تعديل الشراء",
        "delete_purchase": "🗑️ حذف الشراء",
        "clear_purchases_table": "⚠️ تفريغ جدول المشتريات",
        "print_purchases_list": "🖨️ طباعة سجل المشتريات",
        "supplier": "المورد",
        "unit_cost": "تكلفة الوحدة",
        "add_expense": "+ إضافة مصروف",
        "edit_expense": "✏️ تعديل المصروف",
        "delete_expense": "🗑️ حذف المصروف",
        "print_expenses_list": "🖨️ طباعة سجل المصاريف",
        "expense_title": "بيان المصروف",
        "amount": "المبلغ",
        "from_date": "من تاريخ",
        "to_date": "إلى تاريخ",
        "view_report": "عرض التقرير",
        "export_csv": "تصدير CSV",
        "export_excel": "📗 تصدير إكسل",
        "import_customers_excel": "📥 استيراد زبائن (إكسل)",
        "download_customers_template": "📄 نموذج الزبائن",
        "excel_missing": "مكتبة 'openpyxl' غير مثبتة.\nثبّتها مرة واحدة بالأمر: pip install openpyxl",
        "total_sales": "إجمالي المبيعات",
        "total_purchases": "إجمالي المشتريات",
        "cogs": "تكلفة البضاعة المباعة",
        "gross_margin": "هامش الربح الإجمالي",
        "total_expenses": "إجمالي المصاريف",
        "net_profit": "صافي الأرباح التقديرية",
        "users_tab": "إدارة المستخدمين وكلمات المرور",
        "company_tab": "معلومات المؤسسة والطباعة",
        "add_user": "+ إضافة مستخدم جديد",
        "edit_user": "✏️ تعديل المستخدم",
        "change_pwd": "🔑 تغيير كلمة المرور",
        "delete_user": "🗑️ حذف المستخدم",
        "print_users_list": "🖨️ طباعة قائمة المستخدمين",
        "username": "اسم المستخدم",
        "password": "كلمة المرور",
        "confirm_pwd": "تأكيد كلمة المرور",
        "admin": "مسير (Admin)",
        "cashier": "بائع (Caissier)",
        "save_settings": "حفظ الإعدادات",
        "company_name": "اسم المحل / المؤسسة",
        "company_activity": "النشاط التجاري",
        "company_phone": "هاتف المحل",
        "company_address": "العنوان الكامل",
        "rc": "السجل التجاري (RC)",
        "nif": "الرقم الجبائي (NIF)",
        "invoice_format": "نوع الطباعة الافتراضي",
        "format_a4": "فاتورة A4 عادية",
        "format_ticket": "وصل حراري 80mm",
        "invoice_footer_msg": "رسالة أسفل الفاتورة",
        "confirm_delete": "هل أنت متأكد من رغبتك في الحذف؟",
        "saved_success": "تمت العملية بنجاح.",
        "field_required": "يرجى ملء جميع الحقول المطلوبة.",
        "pwd_mismatch": "كلمتا المرور غير متطابقتين.",
        "stock_insufficient": "الكمية المطلوبة أكبر من المخزون المتوفر.",
        "product_not_found": "لم يتم العثور على أي منتج بهذا الباركود.",
        "cannot_delete_self": "لا يمكنك حذف حسابك المسجل به حالياً.",
        "cannot_delete_last_admin": "لا يمكن حذف آخر مسير في النظام.",
        "customer_debt_warning": "هذا الزبون لديه ديون مسجلة. هل أنت متأكد من حذفه وحذف سجلاته؟",
        "sale_delete_warning": "سيتم حذف هذه الفاتورة وإرجاع كميات المنتجات إلى المخزون تلقائياً. هل تريد المتابعة؟",
        "purchase_delete_warning": "سيتم حذف فاتورة الشراء هذه وخصم كمياتها من المخزون تلقائياً. هل تريد المتابعة؟",
    }
}

CURRENT_LANG = "ar"

def t(key):
    return I18N.get(CURRENT_LANG, I18N["fr"]).get(key, key)

def money(v, show_curr=True):
    try:
        val = float(v)
        formatted = f"{val:,.2f}".replace(",", " ")
        if show_curr:
            curr = t("curr")
            return f"{formatted} {curr}"
        return formatted
    except Exception:
        curr = t("curr")
        return f"0.00 {curr}" if show_curr else "0.00"

def num(v):
    try:
        return float(str(v).replace(",", ".").strip())
    except Exception:
        return 0.0

def today():
    return date.today().isoformat()

# ============================================================
# محرك الترجمة الذكي للمنتجات إلى العربية
# ============================================================
import re
LEXICON_AR = [
    (r'\bPATE\s+A\s+TARTINER\b', 'شوكولاتة طلي'),
    (r'\bBARRE\s+CHOCOLAT\b', 'بار شوكولاتة'),
    (r'\bBARRE\b', 'بار'),
    (r'\bPACK\s+FAMILY\b', 'حزمة عائلية'),
    (r'\bFAMILY\b', 'عائلي'),
    (r'\bSANS\s+ENROBAGE\b', 'بدون تغليف'),
    (r'\bENROBE\b|\bENROBEE\b', 'مغلف بالشوكولاتة'),
    (r'\bAU\s+LAIT\b', 'بالحليب'),
    (r'\bCREME\s+VANILLE\b|\bCRÈME\s+VANILLE\b', 'كريمة الفانيليا'),
    (r'\bCREME\s+CARAMEL\b|\bCRÈME\s+CARAMEL\b', 'كريمة الكراميل'),
    (r'\bCREME\s+FRAISE\b|\bCRÈME\s+FRAISE\b', 'كريمة الفراولة'),
    (r'\bCREME\s+NOISETTE\b|\bCRÈME\s+NOISETTE\b', 'كريمة البندق'),
    (r'\bECLAT\s+D\'AMANDES\b|\bECLATS\s+D\'AMANDES\b', 'قطع اللوز'),
    (r'\bECLATS\s+PISTACHE\b|\bECLAT\s+PISTACHE\b', 'قطع الفستق'),
    (r'\bECLAT\s+BISCUIT\b', 'قطع البسكويت'),
    (r'\bFOURRE\s+CREME\b|\bFOURREE\s+CREME\b|\bF.\s+CREME\b|\bF.CREME\b', 'محشو كريمة'),
    (r'\bFOURRE\b|\bFOURRÉE\b|\bFOURREE\b', 'محشو'),
    (r'\bF.\s+PRALINE\b|\bF.PRALINE\b', 'محشو برالين'),
    (r'\bF.\s+PISTACH\b', 'محشو فستق'),
    (r'\bTABLETTE\b|\bTAB\b', 'لوح شوكولاتة'),
    (r'\bBISCUIT\b|\bBISCUITS\b|\bBIS\b', 'بسكويت'),
    (r'\bGAUFRETTE\b|\bGAUFRETTES\b|\bWAFER\b', 'قوفرات'),
    (r'\bBONBON\b|\bBONBONS\b', 'حلوى'),
    (r'\bGALETTE\b|\bGALETTES\b', 'غاليت'),
    (r'\bDONUTS\b|\bDONUT\b', 'دونات'),
    (r'\bGOUTER\b|\bGOÛTER\b', 'لمجة'),
    (r'\bCHOCOLAT\b|\bCHOCOLE\b|\bCHOCO\b', 'شوكولاتة'),
    (r'\bVANILLE\b|\bVAN\b', 'فانيليا'),
    (r'\bCARAMEL\b', 'كراميل'),
    (r'\bFRAISE\b', 'فراولة'),
    (r'\bNOISETTE\b|\bNOISETTES\b', 'بندق'),
    (r'\bENTIER\b', 'كامل'),
    (r'\bAMANDE\b|\bAMANDES\b', 'لوز'),
    (r'\bARACHIDE\b|\bARACHID\b|\bCACAHUETE\b', 'فول سوداني (كاوكاو)'),
    (r'\bPISTACHE\b|\bPISTACH\b', 'فستق'),
    (r'\bLAIT\b', 'حليب'),
    (r'\bBLANC\b|\bBLN\b', 'أبيض'),
    (r'\bNOIR\b', 'أسود'),
    (r'\bCREME\b|\bCRÈME\b', 'كريمة'),
    (r'\bMINI\b', 'ميني'),
    (r'\bROUND\b', 'دائري'),
    (r'\bBREAK\b', 'بريك'),
    (r'\bSACHET\b', 'كيس'),
    (r'\bBOITE\b|\bBOÎTE\b', 'علبة'),
    (r'\bCARTON\b', 'كرتون'),
    (r'\bNEW\b', 'جديد'),
    (r'\bPCS\b|\bPCES\b', 'قطعة'),
    (r'\bCS\b', 'علبة'),
    (r'(\d+)\s*C\b', r'\1 كرتون'),
    (r'(\d+)\s*G\b|(\d+)\s*GR\b', r'\1 غرام'),
    (r'(\d+)\s*KG\b', r'\1 كغ'),
    (r'\bMAXON\b', 'ماكسون'),
    (r'\bBIMO\b', 'بيمو'),
    (r'\bKOOL\b', 'كول'),
    (r'\bHAPPY\b', 'هابي'),
    (r'\bMOMENT\b', 'مومنت'),
    (r'\bBLAXON\b', 'بلاكسون'),
    (r'\bGIANDUITTO\b', 'جياندويتو'),
    (r'\bTWIST\b', 'تويست'),
    (r'\bWINNERS\b|\bWINNER\b', 'وينرز'),
    (r'\bDUBAI\b', 'دبي'),
    (r'\bBRUXELLES\b', 'بروكسل'),
    (r'\bGOLD\b', 'ذهبي'),
    (r'\bORIGINAL\b', 'أصلي'),
    (r'\bBIBILA\b', 'بيبيلا'),
]

def is_arabic_text(text):
    return any('\u0600' <= char <= '\u06FF' for char in str(text))

def translate_product_to_arabic(text):
    if not text: return ""
    text_str = str(text).strip()
    if is_arabic_text(text_str):
        return text_str
    res = text_str
    for pattern, repl in LEXICON_AR:
        res = re.sub(pattern, repl, res, flags=re.IGNORECASE)
    res = re.sub(r'\s+', ' ', res).strip()
    return res

def display_product_name(text):
    if not text: return ""
    if CURRENT_LANG == "ar":
        return translate_product_to_arabic(text)
    return str(text)

# ============================================================
# طبقة قاعدة البيانات SQLite
# ============================================================
class StoreDB:
    def __init__(self, path):
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys=ON")
        self.init_db()

    def init_db(self):
        cur = self.conn.cursor()
        cur.executescript("""
        CREATE TABLE IF NOT EXISTS users(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT DEFAULT 'Admin',
            created_at TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS settings(
            key TEXT PRIMARY KEY,
            value TEXT
        );
        CREATE TABLE IF NOT EXISTS products(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            category TEXT DEFAULT '',
            barcode TEXT DEFAULT '',
            charg REAL DEFAULT 1,
            purchase_price REAL DEFAULT 0,
            sale_price REAL DEFAULT 0,
            stock REAL DEFAULT 0,
            min_stock REAL DEFAULT 0,
            unit TEXT DEFAULT 'Boîte',
            expiry TEXT DEFAULT '',
            created_at TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS customers(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            phone TEXT DEFAULT '',
            address TEXT DEFAULT '',
            notes TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS sales(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sale_date TEXT NOT NULL,
            customer_id INTEGER,
            total REAL NOT NULL,
            paid REAL DEFAULT 0,
            payment_method TEXT DEFAULT 'Espèces',
            notes TEXT DEFAULT '',
            FOREIGN KEY(customer_id) REFERENCES customers(id) ON DELETE SET NULL
        );
        CREATE TABLE IF NOT EXISTS sale_items(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sale_id INTEGER NOT NULL,
            product_id INTEGER,
            product_name TEXT NOT NULL,
            charg REAL DEFAULT 1,
            quantity REAL NOT NULL,
            unit_price REAL NOT NULL,
            unit_cost REAL DEFAULT 0,
            subtotal REAL NOT NULL,
            FOREIGN KEY(sale_id) REFERENCES sales(id) ON DELETE CASCADE,
            FOREIGN KEY(product_id) REFERENCES products(id) ON DELETE SET NULL
        );
        CREATE TABLE IF NOT EXISTS purchases(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            purchase_date TEXT NOT NULL,
            supplier TEXT DEFAULT '',
            total REAL NOT NULL,
            paid REAL DEFAULT 0,
            notes TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS purchase_items(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            purchase_id INTEGER NOT NULL,
            product_id INTEGER,
            product_name TEXT NOT NULL,
            quantity REAL NOT NULL,
            unit_cost REAL NOT NULL,
            subtotal REAL NOT NULL,
            FOREIGN KEY(purchase_id) REFERENCES purchases(id) ON DELETE CASCADE,
            FOREIGN KEY(product_id) REFERENCES products(id) ON DELETE SET NULL
        );
        CREATE TABLE IF NOT EXISTS expenses(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            expense_date TEXT NOT NULL,
            title TEXT NOT NULL,
            amount REAL NOT NULL,
            notes TEXT DEFAULT ''
        );
        """)
        self.conn.commit()

        # Migrations
        user_cols = [c[1] for c in cur.execute("PRAGMA table_info(users)").fetchall()]
        if "role" not in user_cols:
            cur.execute("ALTER TABLE users ADD COLUMN role TEXT DEFAULT 'Admin'")
        if "created_at" not in user_cols:
            cur.execute("ALTER TABLE users ADD COLUMN created_at TEXT DEFAULT ''")
        prod_cols = [c[1] for c in cur.execute("PRAGMA table_info(products)").fetchall()]
        if "charg" not in prod_cols:
            cur.execute("ALTER TABLE products ADD COLUMN charg REAL DEFAULT 1")
        sale_item_cols = [c[1] for c in cur.execute("PRAGMA table_info(sale_items)").fetchall()]
        if "charg" not in sale_item_cols:
            cur.execute("ALTER TABLE sale_items ADD COLUMN charg REAL DEFAULT 1")
        self.conn.commit()

        if not self.one("SELECT id FROM users LIMIT 1"):
            self.run("INSERT INTO users (username, password, role, created_at) VALUES (?, ?, ?, ?)",
                     ('admin', 'admin123', 'Admin', today()))

        default_settings = {
            "app_language": "ar",
            "company_name": "Trafik Saveurs",
            "company_activity": "Gros en Biscuiterie & Confiserie",
            "company_phone": "0561 88 45 06",
            "company_address": "Alger, Algérie",
            "company_rc": "",
            "company_nif": "",
            "invoice_format": "A4",
            "invoice_footer": "شكراً لتعاملكم معنا ونتطلع لخدمتكم دائماً • Merci de votre confiance"
        }
        for k, v in default_settings.items():
            if not self.one("SELECT key FROM settings WHERE key=?", (k,)):
                self.run("INSERT INTO settings (key, value) VALUES (?, ?)", (k, v))

        for tbl in ["products", "customers", "users", "expenses", "sales", "purchases"]:
            self.resequence_table(tbl)

    def all(self, sql, args=()):
        return self.conn.execute(sql, args).fetchall()

    def one(self, sql, args=()):
        return self.conn.execute(sql, args).fetchone()

    def run(self, sql, args=()):
        cur = self.conn.execute(sql, args)
        self.conn.commit()
        return cur.lastrowid

    def get_setting(self, key, default=""):
        r = self.one("SELECT value FROM settings WHERE key=?", (key,))
        return r["value"] if r else default

    def set_setting(self, key, value):
        self.run("INSERT INTO settings (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                 (key, str(value)))

    def resequence_table(self, table_name, id_col="id"):
        cur = self.conn.cursor()
        cur.execute("PRAGMA foreign_keys = OFF")
        try:
            rows = cur.execute(f"SELECT {id_col} FROM {table_name} ORDER BY {id_col} ASC").fetchall()
            if not rows:
                cur.execute("UPDATE sqlite_sequence SET seq=0 WHERE name=?", (table_name,))
                if cur.rowcount == 0:
                    cur.execute("INSERT INTO sqlite_sequence (name, seq) VALUES (?, 0)", (table_name,))
                self.conn.commit()
                return
            for new_id, r in enumerate(rows, start=1):
                old_id = r[id_col]
                if old_id != new_id:
                    if table_name == "products":
                        cur.execute("UPDATE sale_items SET product_id=? WHERE product_id=?", (new_id, old_id))
                        cur.execute("UPDATE purchase_items SET product_id=? WHERE product_id=?", (new_id, old_id))
                    elif table_name == "customers":
                        cur.execute("UPDATE sales SET customer_id=? WHERE customer_id=?", (new_id, old_id))
                    elif table_name == "sales":
                        cur.execute("UPDATE sale_items SET sale_id=? WHERE sale_id=?", (new_id, old_id))
                    elif table_name == "purchases":
                        cur.execute("UPDATE purchase_items SET purchase_id=? WHERE purchase_id=?", (new_id, old_id))
                    cur.execute(f"UPDATE {table_name} SET {id_col}=? WHERE {id_col}=?", (new_id, old_id))
            max_id = len(rows)
            cur.execute("UPDATE sqlite_sequence SET seq=? WHERE name=?", (max_id, table_name))
            if cur.rowcount == 0:
                cur.execute("INSERT INTO sqlite_sequence (name, seq) VALUES (?, ?)", (table_name, max_id))
            self.conn.commit()
        finally:
            cur.execute("PRAGMA foreign_keys = ON")

    def close(self):
        try: self.conn.close()
        except Exception: pass

# ============================================================
# شاشة التحميل العصرية (Splash Screen) - تحاكي تصميم HTML
# ============================================================
class SplashScreen(tk.Tk):
    """شاشة تحميل عصرية تحاكي تصميم HTML المرفق"""
    def __init__(self):
        super().__init__()
        self.title("Trafik Saveurs - Chargement...")
        self.configure(bg=HEADER_BG)
        self.resizable(False, False)
        self.overrideredirect(True)  # إخفاء إطار النافذة
        apply_window_icon(self)

        # الأبعاد
        self.width = 600
        self.height = 450
        center_window(self, self.width, self.height)

        self.progress = 0
        self.current_lang = "ar"
        self.build_ui()
        self.start_loading()

    def build_ui(self):
        # Canvas للرسم
        self.canvas = tk.Canvas(self, width=self.width, height=self.height,
                                bg=HEADER_BG, highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        # رسم الدوائر الزخرفية (مثل HTML)
        self.draw_decorations()

        # الشعار
        self.canvas.create_text(self.width//2, 80,
                                text="🍪", font=("Segoe UI", 50), fill=GOLD)

        # اسم التطبيق
        self.canvas.create_text(self.width//2, 140,
                                text="Trafik Saveurs",
                                font=("Segoe UI", 28, "bold"), fill=GOLD)

        # الوصف
        self.desc_text = self.canvas.create_text(self.width//2, 175,
                                text="نظام تسيير تجارة الحلويات والبسكويت بالجملة",
                                font=("Segoe UI", 12), fill="#E7DDD5")

        # الساعة الرقمية
        self.time_text = self.canvas.create_text(self.width//2, 220,
                                text="--:--:--",
                                font=("Segoe UI", 22, "bold"), fill=GOLD)

        # زر تبديل اللغة
        self.lang_btn_id = self.canvas.create_rectangle(
            self.width//2 - 60, 250, self.width//2 + 60, 280,
            fill=GOLD, outline="", tags="lang_btn")
        self.canvas.create_text(self.width//2, 265,
                                text="Français",
                                font=("Segoe UI", 10, "bold"), fill=BROWN, tags="lang_btn")
        self.canvas.tag_bind("lang_btn", "<Button-1>", self.toggle_lang)

        # شريط التقدم
        self.canvas.create_rectangle(50, 320, self.width-50, 335,
                                     fill="#1A0A06", outline="", tags="progress_bg")
        self.progress_bar = self.canvas.create_rectangle(50, 320, 50, 335,
                                     fill=GOLD, outline="", tags="progress_bar")

        # نص التحميل
        self.loading_text = self.canvas.create_text(self.width//2, 360,
                                text="جاري التحميل...",
                                font=("Segoe UI", 10), fill="#E7DDD5")

        # الإصدار
        self.canvas.create_text(self.width//2, 410,
                                text="Version 3.0 • © 2026 Trafik Saveurs",
                                font=("Segoe UI", 8), fill=MUTED)

        # ربط النقر للإغلاق
        self.canvas.bind("<Button-1>", lambda e: None)

    def draw_decorations(self):
        """رسم الدوائر الزخرفية مثل HTML"""
        colors = [GOLD, BROWN, "#C78800", "#542B1A"]
        for i in range(12):
            angle = (i * 30 - 90) * 3.14159 / 180
            x = self.width//2 + int(250 * __import__('math').cos(angle))
            y = self.height//2 + int(200 * __import__('math').sin(angle))
            size = 8 + (i % 3) * 4
            self.canvas.create_oval(x-size, y-size, x+size, y+size,
                                    fill=colors[i % len(colors)], outline="",
                                    tags="deco")

    def toggle_lang(self, event=None):
        self.current_lang = "fr" if self.current_lang == "ar" else "ar"
        global CURRENT_LANG
        CURRENT_LANG = self.current_lang

        if self.current_lang == "ar":
            self.canvas.itemconfig(self.desc_text, text="نظام تسيير تجارة الحلويات والبسكويت بالجملة")
            self.canvas.itemconfig(self.loading_text, text="جاري التحميل...")
            self.canvas.itemconfig(self.lang_btn_id, text="Français")
            # تحديث النص على الزر
            self.canvas.delete("lang_text")
            self.canvas.create_text(self.width//2, 265,
                                    text="Français",
                                    font=("Segoe UI", 10, "bold"), fill=BROWN, tags="lang_text")
        else:
            self.canvas.itemconfig(self.desc_text, text="Système de Gestion de Biscuiterie & Confiserie en Gros")
            self.canvas.itemconfig(self.loading_text, text="Chargement en cours...")
            self.canvas.delete("lang_text")
            self.canvas.create_text(self.width//2, 265,
                                    text="العربية",
                                    font=("Segoe UI", 10, "bold"), fill=BROWN, tags="lang_text")

    def update_clock(self):
        """تحديث الساعة الرقمية"""
        now = datetime.now()
        time_str = now.strftime("%H:%M:%S")
        self.canvas.itemconfig(self.time_text, text=time_str)

    def start_loading(self):
        """بدء عملية التحميل"""
        self.loading_steps = [
            (10, "تحميل قاعدة البيانات..."),
            (25, "Chargement de la base de données..."),
            (40, "تهيئة الواجهة..."),
            (55, "Initialisation de l'interface..."),
            (70, "تحميل المنتجات..."),
            (85, "Chargement des produits..."),
            (95, "الاستعداد للتشغيل..."),
            (100, "Prêt à démarrer...")
        ]
        self.step_index = 0
        self.animate_loading()
        self.update_clock_loop()

    def update_clock_loop(self):
        self.update_clock()
        self.after(1000, self.update_clock_loop)

    def animate_loading(self):
        if self.step_index < len(self.loading_steps):
            progress, text = self.loading_steps[self.step_index]
            self.progress = progress

            # تحديث شريط التقدم
            bar_width = int((self.width - 100) * (progress / 100))
            self.canvas.coords(self.progress_bar, 50, 320, 50 + bar_width, 335)

            # تحديث نص التحميل
            self.canvas.itemconfig(self.loading_text, text=text)

            self.step_index += 1
            self.after(400, self.animate_loading)
        else:
            # انتهاء التحميل - الانتقال للتطبيق الرئيسي
            self.after(500, self.open_main_app)

    def open_main_app(self):
        self.destroy()
        app = LoginApp()
        app.mainloop()

# ============================================================
# شاشة تسجيل الدخول
# ============================================================
class LoginApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Trafik Saveurs - Connexion / تسجيل الدخول")
        center_window(self, 480, 560)
        self.configure(bg=BG)
        self.resizable(False, False)
        apply_window_icon(self)

        self.db = StoreDB(DB_PATH)
        global CURRENT_LANG
        CURRENT_LANG = self.db.get_setting("app_language", "ar")
        self.build_ui()

    def toggle_lang(self):
        global CURRENT_LANG
        CURRENT_LANG = "fr" if CURRENT_LANG == "ar" else "ar"
        self.db.set_setting("app_language", CURRENT_LANG)
        for w in self.winfo_children(): w.destroy()
        self.build_ui()

    def build_ui(self):
        top_bar = tk.Frame(self, bg=HEADER_BG, height=115)
        top_bar.pack(fill="x")
        top_bar.pack_propagate(False)

        lang_btn = tk.Button(top_bar, text="Français" if CURRENT_LANG == "ar" else "العربية",
                             command=self.toggle_lang, bg=GOLD, fg=BROWN,
                             font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2", padx=10, pady=3)
        lang_btn.pack(side="right", padx=15, pady=10)

        header_center = tk.Frame(top_bar, bg=HEADER_BG)
        header_center.pack(expand=True, fill="both")

        titles_wrap = tk.Frame(header_center, bg=HEADER_BG)
        titles_wrap.place(relx=0.5, rely=0.5, anchor="center")

        if LOGO_PATH.exists() and PIL_AVAILABLE:
            try:
                logo_im = Image.open(LOGO_PATH).resize((50, 50), Image.Resampling.LANCZOS)
                self._login_logo = ImageTk.PhotoImage(logo_im)
                tk.Label(titles_wrap, image=self._login_logo, bg=HEADER_BG).pack(side="left", padx=(0, 10))
            except Exception:
                pass

        titles_box = tk.Frame(titles_wrap, bg=HEADER_BG)
        titles_box.pack(side="left")
        tk.Label(titles_box, text="Trafik Saveurs", bg=HEADER_BG, fg=GOLD,
                 font=("Segoe UI", 18, "bold")).pack(anchor="w")
        tk.Label(titles_box, text=t("subtitle"), bg=HEADER_BG, fg="#E7DDD5",
                 font=("Segoe UI", 8)).pack(anchor="w")

        card = tk.Frame(self, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
        card.place(relx=0.5, rely=0.58, anchor="center", width=380, height=330)

        tk.Label(card, text="تسجيل الدخول" if CURRENT_LANG=="ar" else "Connexion",
                 bg=PANEL, fg=BROWN, font=("Segoe UI", 15, "bold")).pack(pady=(20, 15))

        tk.Label(card, text=t("username"), bg=PANEL, fg=BROWN,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=35)
        self.entry_user = tk.Entry(card, font=("Segoe UI", 11), relief="solid", bd=1)
        self.entry_user.pack(fill="x", padx=35, pady=(4, 12), ipady=5)
        self.entry_user.insert(0, "admin")

        tk.Label(card, text=t("password"), bg=PANEL, fg=BROWN,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=35)
        self.entry_pass = tk.Entry(card, font=("Segoe UI", 11), show="•", relief="solid", bd=1)
        self.entry_pass.pack(fill="x", padx=35, pady=(4, 18), ipady=5)
        self.entry_pass.insert(0, "admin123")
        self.entry_pass.bind("<Return>", lambda e: self.check_login())

        btn_login = tk.Button(card, text="دخول" if CURRENT_LANG=="ar" else "Se Connecter",
                              command=self.check_login, bg=GOLD, fg=BROWN,
                              font=("Segoe UI", 11, "bold"), relief="flat", cursor="hand2",
                              activebackground=GOLD_HOVER)
        btn_login.pack(fill="x", padx=35, ipady=8)

        tk.Label(self, text="نظام مكتبي محلي • Trafik Saveurs 2026", bg=BG, fg=MUTED,
                 font=("Segoe UI", 8)).pack(side="bottom", pady=10)

    def check_login(self):
        u = self.entry_user.get().strip()
        p = self.entry_pass.get().strip()
        user = self.db.one("SELECT * FROM users WHERE username=? AND password=?", (u, p))
        if user:
            user_data = dict(user)
            self.db.close()
            self.destroy()
            app = App(user_data)
            app.mainloop()
        else:
            beep_err()
            err_msg = "اسم المستخدم أو كلمة المرور غير صحيحة" if CURRENT_LANG=="ar" else "Nom d'utilisateur ou mot de passe incorrect"
            messagebox.showerror("خطأ" if CURRENT_LANG=="ar" else "Erreur", err_msg, parent=self)

# ============================================================
# التطبيق الرئيسي الكامل
# ============================================================
class App(tk.Tk):
    def __init__(self, current_user):
        super().__init__()
        self.current_user = current_user
        self.db = StoreDB(DB_PATH)
        global CURRENT_LANG
        CURRENT_LANG = self.db.get_setting("app_language", "ar")
        self.title(t("app_title"))

        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        app_w = min(1280, int(sw * 0.96))
        app_h = min(730, int(sh * 0.86))
        self.minsize(800, 480)
        center_window(self, app_w, app_h)

        if sw <= 1366 or sh <= 768:
            try: self.state('zoomed')
            except Exception: pass

        self.configure(bg=BG)
        if ICON_PATH.exists():
            try: self.iconbitmap(str(ICON_PATH))
            except Exception: pass

        self.style_app()
        self.build_shell()
        self.show_dashboard()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def style_app(self):
        s = ttk.Style(self)
        try: s.theme_use("clam")
        except Exception: pass
        s.configure("Treeview", font=("Segoe UI", 9), rowheight=27, background="white",
                    fieldbackground="white", foreground=BROWN, bordercolor=BORDER)
        s.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"),
                    background="#EDE4D6", foreground=BROWN, relief="flat")
        s.map("Treeview", background=[("selected", "#FBEABF")], foreground=[("selected", BROWN)])
        s.configure("TCombobox", padding=4, font=("Segoe UI", 9))
        s.configure("TNotebook", background=BG, borderwidth=0)
        s.configure("TNotebook.Tab", font=("Segoe UI", 9, "bold"), padding=(12, 6))

    def switch_language(self):
        global CURRENT_LANG
        CURRENT_LANG = "fr" if CURRENT_LANG == "ar" else "ar"
        self.db.set_setting("app_language", CURRENT_LANG)
        self.title(t("app_title"))
        for w in self.winfo_children(): w.destroy()
        self.style_app()
        self.build_shell()
        self.show_dashboard()

    def build_shell(self):
        header = tk.Frame(self, bg=HEADER_BG, height=75)
        header.pack(fill="x")
        header.pack_propagate(False)

        left_box = tk.Frame(header, bg=HEADER_BG)
        left_box.pack(side="left", padx=15, pady=8)

        if LOGO_PATH.exists() and PIL_AVAILABLE:
            try:
                logo_im = Image.open(LOGO_PATH).resize((50, 50), Image.Resampling.LANCZOS)
                self._header_logo = ImageTk.PhotoImage(logo_im)
                tk.Label(left_box, image=self._header_logo, bg=HEADER_BG).pack(side="left", padx=(0, 10))
            except Exception:
                pass

        titles_box = tk.Frame(left_box, bg=HEADER_BG)
        titles_box.pack(side="left")
        comp_name = self.db.get_setting("company_name", APP_NAME)
        tk.Label(titles_box, text=comp_name, bg=HEADER_BG, fg=GOLD,
                 font=("Segoe UI", 18, "bold")).pack(anchor="w")
        comp_act = self.db.get_setting("company_activity", t("subtitle"))
        tk.Label(titles_box, text=comp_act, bg=HEADER_BG, fg="#F0E8E1",
                 font=("Segoe UI", 9)).pack(anchor="w")

        right_box = tk.Frame(header, bg=HEADER_BG)
        right_box.pack(side="right", padx=15, pady=10)
        u_role = t("admin") if self.current_user.get("role") == "Admin" else t("cashier")
        user_badge = f"👤 {self.current_user.get('username')} ({u_role})"
        tk.Label(right_box, text=user_badge, bg="#3E2018", fg=GOLD,
                 font=("Segoe UI", 9, "bold"), padx=8, pady=3).pack(side="left", padx=6)
        lang_label = "Français 🌐" if CURRENT_LANG == "ar" else "العربية 🌐"
        tk.Button(right_box, text=lang_label, command=self.switch_language,
                  bg=GOLD, fg=BROWN, font=("Segoe UI", 9, "bold"),
                  relief="flat", cursor="hand2", padx=8, pady=3).pack(side="left")

        self.nav = tk.Frame(self, bg="#E8DFD1", height=44)
        self.nav.pack(fill="x")
        nav_items = [
            ("dashboard", self.show_dashboard),
            ("sales", self.show_sales),
            ("products", self.show_products),
            ("customers", self.show_customers),
            ("purchases", self.show_purchases),
            ("expenses", self.show_expenses),
            ("reports", self.show_reports),
            ("settings", self.show_settings),
            ("backup", self.backup),
            ("logout", self.logout),
        ]
        for key, cmd in nav_items:
            bg_c = RED if key == "logout" else "#E8DFD1"
            fg_c = "white" if key == "logout" else BROWN
            btn = tk.Button(self.nav, text=t(key), command=cmd, bg=bg_c, fg=fg_c,
                            activebackground=GOLD, relief="flat", font=("Segoe UI", 9, "bold"),
                            padx=11, pady=7, cursor="hand2")
            btn.pack(side="left" if CURRENT_LANG=="fr" else "right")

        self.body = tk.Frame(self, bg=BG)
        self.body.pack(fill="both", expand=True, padx=15, pady=10)

        self.status = tk.Label(self, text=f"{t('offline_badge')}  |  SQLite: {DB_PATH.name}",
                               anchor="w" if CURRENT_LANG=="fr" else "e",
                               bg="#E4DBCB", fg=MUTED, font=("Segoe UI", 8), padx=12, pady=3)
        self.status.pack(fill="x", side="bottom")

    def clear(self):
        for w in self.body.winfo_children(): w.destroy()

    def heading(self, title_text, subtitle_text=""):
        top = tk.Frame(self.body, bg=BG)
        top.pack(fill="x", pady=(0, 10))
        anchor_val = "w" if CURRENT_LANG=="fr" else "e"
        tk.Label(top, text=title_text, bg=BG, fg=BROWN, font=("Segoe UI", 16, "bold")).pack(anchor=anchor_val)
        if subtitle_text:
            tk.Label(top, text=subtitle_text, bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(anchor=anchor_val, pady=(1, 0))

    def stat_card(self, parent, title, value, color=GOLD, alert=False):
        f = tk.Frame(parent, bg=PANEL, highlightbackground=color if alert else BORDER,
                     highlightthickness=2 if alert else 1)
        tk.Frame(f, bg=color, height=4).pack(fill="x")
        anchor_val = "w" if CURRENT_LANG=="fr" else "e"
        tk.Label(f, text=title, bg=PANEL, fg=MUTED, font=("Segoe UI", 9, "bold")).pack(anchor=anchor_val, padx=10, pady=(8, 2))
        tk.Label(f, text=value, bg=PANEL, fg=color if alert else BROWN,
                 font=("Segoe UI", 14, "bold")).pack(anchor=anchor_val, padx=10, pady=(0, 10))
        return f

    def create_tree(self, parent, columns, headings, widths=None, height=10):
        wrap = tk.Frame(parent, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
        wrap.pack(fill="both", expand=True)
        tree = ttk.Treeview(wrap, columns=columns, show="headings", height=height)
        if CURRENT_LANG == "ar":
            tree.configure(displaycolumns=list(reversed(columns)))
            anchor_dir = "e"
        else:
            tree.configure(displaycolumns=list(columns))
            anchor_dir = "w"
        for i, (c, h) in enumerate(zip(columns, headings)):
            tree.heading(c, text=h, anchor=anchor_dir)
            w = widths[i] if widths and i < len(widths) else 120
            tree.column(c, width=w, anchor=anchor_dir)
        y = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        x = ttk.Scrollbar(wrap, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=y.set, xscrollcommand=x.set)
        tree.grid(row=0, column=0, sticky="nsew")
        y.grid(row=0, column=1, sticky="ns")
        x.grid(row=1, column=0, sticky="ew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        return tree

    def print_table_report(self, title, headers, rows, filename_prefix="rapport"):
        comp_name = html.escape(self.db.get_setting("company_name", APP_NAME))
        comp_act = html.escape(self.db.get_setting("company_activity", t("subtitle")))
        comp_phone = html.escape(self.db.get_setting("company_phone", "0561 88 45 06"))
        comp_addr = html.escape(self.db.get_setting("company_address", ""))
        th_html = "".join(f"<th>{html.escape(h)}</th>" for h in headers)
        tr_html = ""
        for r in rows:
            tds = "".join(f"<td>{html.escape(str(c))}</td>" for c in r)
            tr_html += f"<tr>{tds}</tr>"
        doc = f"""<!doctype html>
        <html lang="fr" dir="ltr">
        <head><meta charset="utf-8"><title>{title}</title>
        <style>
            body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 25px; color: #2B160F; }}
            .header {{ border-bottom: 3px solid #E09C00; padding-bottom: 10px; margin-bottom: 15px; }}
            .header h1 {{ margin: 0; color: #3B1B12; font-size: 22px; }}
            .header p {{ margin: 3px 0; color: #555; font-size: 12px; }}
            h2 {{ color: #E09C00; margin-top: 10px; font-size: 16px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 12px; }}
            th {{ background: #3B1B12; color: #FFF; padding: 8px; font-size: 12px; border: 1px solid #3B1B12; }}
            td {{ border: 1px solid #E0D7CB; padding: 7px 8px; font-size: 11px; text-align: center; }}
            tr:nth-child(even) {{ background: #FAF7F2; }}
            .footer {{ text-align: center; margin-top: 25px; font-size: 11px; color: #777; border-top: 1px solid #DDD; padding-top: 12px; }}
            @media print {{ button {{ display: none; }} body {{ margin: 10mm; }} }}
        </style></head>
        <body>
            <div class="header">
                <h1>{comp_name}</h1>
                <p><b>{comp_act}</b> | Tél: {comp_phone} | {comp_addr}</p>
                <p><b>Date:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
            </div>
            <h2>{html.escape(title)}</h2>
            <table><thead><tr>{th_html}</tr></thead><tbody>{tr_html}</tbody></table>
            <div class="footer">
                <button onclick="window.print()" style="background:#E09C00;color:#3B1B12;border:none;padding:8px 20px;font-weight:bold;font-size:13px;cursor:pointer;border-radius:4px;">🖨️ Imprimer</button>
            </div>
        </body></html>"""
        path = ROOT / f"{filename_prefix}_{today()}.html"
        path.write_text(doc, encoding="utf-8")
        webbrowser.open(path.as_uri())

    # ==========================================
    # 1. لوحة التحكم
    # ==========================================
    def show_dashboard(self):
        self.clear()
        self.heading(t("dashboard"), t("subtitle"))
        d = today()
        sales = self.db.one("SELECT COALESCE(SUM(total),0) n FROM sales WHERE sale_date=?", (d,))["n"]
        paid = self.db.one("SELECT COALESCE(SUM(paid),0) n FROM sales WHERE sale_date=?", (d,))["n"]
        stock_count = self.db.one("SELECT COUNT(*) n FROM products")["n"]
        low_stock = self.db.one("SELECT COUNT(*) n FROM products WHERE stock<=min_stock")["n"]
        debt = self.db.one("SELECT COALESCE(SUM(total-paid),0) n FROM sales")["n"]
        limit_date = (date.today() + timedelta(days=30)).isoformat()
        exp_count = self.db.one("SELECT COUNT(*) n FROM products WHERE expiry != '' AND expiry <= ?", (limit_date,))["n"]

        cards = tk.Frame(self.body, bg=BG)
        cards.pack(fill="x", pady=(0, 10))
        metrics = [
            (t("today_sales"), money(sales), GOLD, False),
            (t("today_paid"), money(paid), GREEN, False),
            (t("products_count"), str(stock_count), BLUE, False),
            (t("critical_stock"), str(low_stock), RED, low_stock > 0),
            (t("expiring_soon"), str(exp_count), ORANGE, exp_count > 0),
            (t("customer_debts"), money(debt), RED, debt > 0),
        ]
        side_align = "left" if CURRENT_LANG=="fr" else "right"
        for title, val, col, alert in metrics:
            self.stat_card(cards, title, val, col, alert).pack(side=side_align, fill="both", expand=True, padx=3)

        if low_stock > 0 or exp_count > 0:
            alert_box = tk.LabelFrame(self.body, text=f"⚠️ {t('alert_section')}", bg="#FFF9E6",
                                      fg=RED, font=("Segoe UI", 10, "bold"), padx=8, pady=6)
            alert_box.pack(fill="x", pady=(0, 10))
            txt_items = []
            if low_stock > 0:
                crit_prods = self.db.all("SELECT name, stock, min_stock FROM products WHERE stock<=min_stock LIMIT 4")
                crit_names = ", ".join(f"{p['name']} ({p['stock']}/{p['min_stock']})" for p in crit_prods)
                txt_items.append(f"• {t('critical_stock')}: {crit_names}")
            if exp_count > 0:
                exp_prods = self.db.all("SELECT name, expiry FROM products WHERE expiry != '' AND expiry <= ? LIMIT 4", (limit_date,))
                exp_names = ", ".join(f"{p['name']} ({p['expiry']})" for p in exp_prods)
                txt_items.append(f"• {t('expiring_soon')}: {exp_names}")
            for it in txt_items:
                tk.Label(alert_box, text=it, bg="#FFF9E6", fg="#8A4800", font=("Segoe UI", 9, "bold")).pack(anchor="w")

        tk.Label(self.body, text=t("recent_sales"), bg=BG, fg=BROWN,
                 font=("Segoe UI", 12, "bold")).pack(anchor="w" if CURRENT_LANG=="fr" else "e", pady=(3, 5))
        t_tree = self.create_tree(self.body, ("id", "date", "customer", "total", "paid", "due"),
                                  (t("invoice_num"), t("date"), t("customer"), t("total"), t("paid"), t("due")),
                                  [70, 110, 240, 130, 130, 130], 6)
        rows = self.db.all("""SELECT s.id, s.sale_date, COALESCE(c.name, ?) customer, s.total, s.paid
                              FROM sales s LEFT JOIN customers c ON c.id=s.customer_id
                              ORDER BY s.id DESC LIMIT 8""", (t("counter_customer"),))
        for r in rows:
            t_tree.insert("", "end", values=(r["id"], r["sale_date"], r["customer"],
                                            money(r["total"]), money(r["paid"]), money(r["total"] - r["paid"])))

    # ==========================================
    # 2. المنتجات والمخزون
    # ==========================================
    def show_products(self):
        self.clear()
        self.heading(t("products"), t("subtitle"))
        bar = tk.Frame(self.body, bg=BG)
        bar.pack(fill="x", pady=(0, 8))
        side = "left" if CURRENT_LANG=="fr" else "right"
        tk.Button(bar, text=t("add_product"), command=self.add_product, bg=GOLD, fg=BROWN,
                  font=("Segoe UI", 9, "bold"), relief="flat", padx=12, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("import_excel"), command=self.import_excel_modal, bg="#E8DFD1", fg=BROWN,
                  font=("Segoe UI", 9, "bold"), relief="flat", padx=11, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("edit_product"), command=self.edit_product, bg="#E8DFD1", fg=BROWN,
                  font=("Segoe UI", 9), relief="flat", padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("delete_product"), command=self.delete_product, bg="#F7D8D5", fg=RED,
                  font=("Segoe UI", 9), relief="flat", padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("print_products"), command=self.print_products_report, bg="#E8DFD1", fg=BROWN,
                  font=("Segoe UI", 9), relief="flat", padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("export_excel"), command=self.export_products_excel, bg="#217A4B", fg="white",
                  font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)

        search_frame = tk.Frame(bar, bg=BG)
        search_frame.pack(side="right" if CURRENT_LANG=="fr" else "left", padx=5)
        tk.Label(search_frame, text=f"🔍 {t('search')}:", bg=BG, font=("Segoe UI", 9, "bold")).pack(side="left", padx=3)
        self.prod_search = tk.Entry(search_frame, font=("Segoe UI", 9), width=18)
        self.prod_search.pack(side="left", padx=3, ipady=3)
        self.prod_search.bind("<KeyRelease>", lambda e: self.filter_products())

        cols = ("id", "name", "category", "barcode", "charg", "buy", "sell", "stock", "unit", "min", "expiry")
        heads = ("N°", t("designation"), t("category"), t("barcode"), t("charg"), t("purchase_price"),
                 t("sale_price"), t("stock"), t("unit"), t("alert_threshold"), t("expiry_date"))
        self.product_tree = self.create_tree(self.body, cols, heads, [45, 220, 110, 120, 65, 95, 95, 65, 65, 75, 95], 9)
        self.product_tree.bind("<Double-1>", lambda e: self.edit_product())

        summary_bar = tk.Frame(self.body, bg="#EDE4D6", height=32, highlightbackground=BORDER, highlightthickness=1)
        summary_bar.pack(fill="x", pady=(6, 0))
        self.prod_summary_label = tk.Label(summary_bar, text="", bg="#EDE4D6", fg=BROWN, font=("Segoe UI", 9, "bold"))
        self.prod_summary_label.pack(side="left" if CURRENT_LANG=="fr" else "right", padx=15, pady=4)
        self.refresh_products()

    def refresh_products(self, query=""):
        if not hasattr(self, "product_tree") or not self.product_tree.winfo_exists(): return
        self.product_tree.delete(*self.product_tree.get_children())
        if query:
            q = f"%{query.strip()}%"
            rows = self.db.all("""SELECT * FROM products WHERE name LIKE ? OR category LIKE ? OR barcode LIKE ?
                                  ORDER BY id ASC""", (q, q, q))
        else:
            rows = self.db.all("SELECT * FROM products ORDER BY id ASC")

        cat_map_ar = {"biscuits": "بسكويت", "biscuit": "بسكويت", "gaufrettes": "قوفرات", "gaufrette": "قوفرات",
                      "confiserie": "حلويات وسكاكر", "chocolat": "شوكولاتة", "gateau": "كعك", "chocolaterie": "شوكولاتة"}
        unit_map_ar = {"carton": "كرتون", "boite": "علبة", "boîte": "علبة", "paquet": "علبة/باقة",
                       "sachet": "كيس", "piece": "قطعة", "pièce": "قطعة"}

        for r in rows:
            tag = "critical" if r["stock"] <= r["min_stock"] else "normal"
            charg_v = f"{r['charg']:g}" if r["charg"] is not None else "0"
            stock_v = f"{r['stock']:g}" if r["stock"] is not None else "0"
            min_v = f"{r['min_stock']:g}" if r["min_stock"] is not None else "0"
            p_name = display_product_name(r["name"])
            p_cat = cat_map_ar.get(str(r["category"] or '').strip().lower(), r["category"] or '-') if CURRENT_LANG=="ar" else (r["category"] or '-')
            p_unit = unit_map_ar.get(str(r["unit"] or '').strip().lower(), r["unit"] or '-') if CURRENT_LANG=="ar" else (r["unit"] or '-')
            self.product_tree.insert("", "end", values=(
                r["id"], p_name, p_cat, r["barcode"] or '-', charg_v,
                money(r["purchase_price"]), money(r["sale_price"]),
                stock_v, p_unit, min_v, r["expiry"] or '-'
            ), tags=(tag,))
        self.product_tree.tag_configure("critical", background="#FFECEB")

        if hasattr(self, "prod_summary_label") and self.prod_summary_label.winfo_exists():
            total_items = len(rows)
            total_stock_qty = sum(r["stock"] for r in rows)
            stock_val = sum(r["stock"] * r["purchase_price"] for r in rows)
            lbl = f"📦 إجمالي المنتجات: {total_items} | إجمالي المخزون: {total_stock_qty:g} | القيمة الإجمالية: {money(stock_val)}" if CURRENT_LANG=="ar" else f"📦 Total Produits : {total_items} | Stock Total : {total_stock_qty:g} | Valeur Totale : {money(stock_val)}"
            self.prod_summary_label.config(text=lbl)

    def filter_products(self):
        self.refresh_products(self.prod_search.get())

    def product_dialog(self, row=None):
        win = tk.Toplevel(self)
        win.title(t("edit_product") if row else t("add_product"))
        center_window(win, 640, 390)
        win.configure(bg=BG)

        header_frame = tk.Frame(win, bg=HEADER_BG, height=38)
        header_frame.pack(fill="x")
        header_frame.pack_propagate(False)
        tk.Label(header_frame, text=t("edit_product") if row else t("add_product"),
                 bg=HEADER_BG, fg=GOLD, font=("Segoe UI", 11, "bold")).pack(pady=7)

        bottom_bar = tk.Frame(win, bg="#EDE4D6", height=48, highlightbackground=BORDER, highlightthickness=1)
        bottom_bar.pack(side="bottom", fill="x")
        bottom_bar.pack_propagate(False)

        form = tk.Frame(win, bg=BG)
        form.pack(fill="both", expand=True, padx=20, pady=10)
        entries = {}

        tk.Label(form, text=f"{t('designation')} *:", bg=BG, fg=BROWN, font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="e", padx=4, pady=3)
        e_name = tk.Entry(form, font=("Segoe UI", 10), width=48)
        e_name.grid(row=0, column=1, columnspan=3, sticky="w", padx=4, pady=3, ipady=2)
        if row: e_name.insert(0, str(row["name"] or ""))
        entries["name"] = e_name

        tk.Label(form, text=f"{t('category')}:", bg=BG, fg=BROWN, font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="e", padx=4, pady=3)
        e_cat = tk.Entry(form, font=("Segoe UI", 9), width=18)
        e_cat.grid(row=1, column=1, sticky="w", padx=4, pady=3, ipady=2)
        if row: e_cat.insert(0, str(row["category"] or ""))
        entries["category"] = e_cat

        tk.Label(form, text=f"{t('barcode')}:", bg=BG, fg=BROWN, font=("Segoe UI", 9, "bold")).grid(row=1, column=2, sticky="e", padx=4, pady=3)
        e_code = tk.Entry(form, font=("Segoe UI", 9), width=18)
        e_code.grid(row=1, column=3, sticky="w", padx=4, pady=3, ipady=2)
        if row: e_code.insert(0, str(row["barcode"] or ""))
        entries["barcode"] = e_code

        tk.Label(form, text=f"{t('purchase_price')} (DA):", bg=BG, fg=BROWN, font=("Segoe UI", 9, "bold")).grid(row=2, column=0, sticky="e", padx=4, pady=3)
        e_buy = tk.Entry(form, font=("Segoe UI", 9), width=18)
        e_buy.grid(row=2, column=1, sticky="w", padx=4, pady=3, ipady=2)
        e_buy.insert(0, str(row["purchase_price"] if row else "0"))
        entries["purchase_price"] = e_buy

        tk.Label(form, text=f"{t('sale_price')} (DA):", bg=BG, fg=BROWN, font=("Segoe UI", 9, "bold")).grid(row=2, column=2, sticky="e", padx=4, pady=3)
        e_sell = tk.Entry(form, font=("Segoe UI", 9), width=18)
        e_sell.grid(row=2, column=3, sticky="w", padx=4, pady=3, ipady=2)
        e_sell.insert(0, str(row["sale_price"] if row else "0"))
        entries["sale_price"] = e_sell

        tk.Label(form, text=f"{t('stock')}:", bg=BG, fg=BROWN, font=("Segoe UI", 9, "bold")).grid(row=3, column=0, sticky="e", padx=4, pady=3)
        e_stock = tk.Entry(form, font=("Segoe UI", 9), width=18)
        e_stock.grid(row=3, column=1, sticky="w", padx=4, pady=3, ipady=2)
        e_stock.insert(0, str(row["stock"] if row else "0"))
        entries["stock"] = e_stock

        tk.Label(form, text=f"{t('alert_threshold')}:", bg=BG, fg=BROWN, font=("Segoe UI", 9, "bold")).grid(row=3, column=2, sticky="e", padx=4, pady=3)
        e_min = tk.Entry(form, font=("Segoe UI", 9), width=18)
        e_min.grid(row=3, column=3, sticky="w", padx=4, pady=3, ipady=2)
        e_min.insert(0, str(row["min_stock"] if row else "5"))
        entries["min_stock"] = e_min

        tk.Label(form, text=f"{t('charg')} (Colis):", bg=BG, fg=BROWN, font=("Segoe UI", 9, "bold")).grid(row=4, column=0, sticky="e", padx=4, pady=3)
        e_charg = tk.Entry(form, font=("Segoe UI", 9), width=18)
        e_charg.grid(row=4, column=1, sticky="w", padx=4, pady=3, ipady=2)
        e_charg.insert(0, str(row["charg"] if row else "1"))
        entries["charg"] = e_charg

        tk.Label(form, text=f"{t('unit')}:", bg=BG, fg=BROWN, font=("Segoe UI", 9, "bold")).grid(row=4, column=2, sticky="e", padx=4, pady=3)
        e_unit = ttk.Combobox(form, values=["Boîte", "Carton", "Pièce", "Sachet", "Pack"], width=16)
        e_unit.set(str(row["unit"] if row else "Boîte"))
        e_unit.grid(row=4, column=3, sticky="w", padx=4, pady=3)
        entries["unit"] = e_unit

        tk.Label(form, text=f"{t('expiry_date')} (YYYY-MM-DD):", bg=BG, fg=BROWN, font=("Segoe UI", 9, "bold")).grid(row=5, column=0, sticky="e", padx=4, pady=3)
        e_exp = tk.Entry(form, font=("Segoe UI", 9), width=18)
        e_exp.grid(row=5, column=1, sticky="w", padx=4, pady=3, ipady=2)
        if row: e_exp.insert(0, str(row["expiry"] or ""))
        entries["expiry"] = e_exp

        def save():
            name = entries["name"].get().strip()
            if not name:
                beep_err()
                messagebox.showerror(t("app_title"), t("field_required"), parent=win)
                return
            vals = [
                name, entries["category"].get().strip(), entries["barcode"].get().strip(),
                num(entries["charg"].get() or "1"),
                num(entries["purchase_price"].get()), num(entries["sale_price"].get()),
                num(entries["stock"].get()), num(entries["min_stock"].get()),
                entries["unit"].get().strip() or "Boîte", entries["expiry"].get().strip()
            ]
            if row:
                self.db.run("""UPDATE products SET name=?, category=?, barcode=?, charg=?, purchase_price=?,
                               sale_price=?, stock=?, min_stock=?, unit=?, expiry=? WHERE id=?""", vals + [row["id"]])
            else:
                self.db.run("""INSERT INTO products(name, category, barcode, charg, purchase_price, sale_price,
                               stock, min_stock, unit, expiry, created_at)
                               VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""", vals + [today()])
            beep_ok()
            win.destroy()
            self.refresh_products()

        save_btn_txt = "💾 حفظ البيانات" if CURRENT_LANG=="ar" else "💾 Enregistrer le Produit"
        tk.Button(bottom_bar, text=save_btn_txt, command=save, bg=GOLD, fg=BROWN,
                  relief="flat", font=("Segoe UI", 10, "bold"), padx=25, pady=6, cursor="hand2").pack(pady=7)

    def add_product(self):
        self.product_dialog()

    def selected_product(self):
        sel = self.product_tree.selection()
        if not sel:
            messagebox.showwarning(t("app_title"), "Sélectionnez un produit" if CURRENT_LANG=="fr" else "يرجى تحديد منتج من الجدول")
            return None
        return self.db.one("SELECT * FROM products WHERE id=?", (self.product_tree.item(sel[0], "values")[0],))

    def edit_product(self):
        r = self.selected_product()
        if r: self.product_dialog(r)

    def delete_product(self):
        r = self.selected_product()
        if r and messagebox.askyesno(t("app_title"), t("confirm_delete")):
            self.db.run("DELETE FROM products WHERE id=?", (r["id"],))
            self.db.resequence_table("products")
            beep_ok()
            self.refresh_products()
            messagebox.showinfo(t("app_title"), t("saved_success"))

    def print_products_report(self):
        rows = self.db.all("SELECT id, name, category, barcode, charg, purchase_price, sale_price, stock, unit, expiry FROM products ORDER BY id ASC")
        data = []
        for r in rows:
            data.append([r["id"], r["name"], r["category"], r["barcode"] or "-", r["charg"],
                         money(r["purchase_price"]), money(r["sale_price"]),
                         f"{r['stock']} {r['unit']}", r["expiry"] or "-"])
        headers = ["N°", t("designation"), t("category"), t("barcode"), t("charg"), t("purchase_price"), t("sale_price"), t("stock"), t("expiry_date")]
        self.print_table_report(t("print_products"), headers, data, "liste_produits")

    # ==========================================
    # 3. المبيعات
    # ==========================================
    def show_sales(self):
        self.clear()
        self.heading(t("sales"), t("subtitle"))
        bar = tk.Frame(self.body, bg=BG)
        bar.pack(fill="x", pady=(0, 8))
        side = "left" if CURRENT_LANG=="fr" else "right"
        tk.Button(bar, text=t("new_invoice"), command=self.new_sale_modal, bg=GOLD, fg=BROWN,
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=12, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("edit_sale"), command=self.edit_sale_dialog, bg="#E8DFD1", fg=BROWN,
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("delete_sale"), command=self.delete_sale_action, bg="#F7D8D5", fg=RED,
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("print_a4"), command=lambda: self.confirm_and_print_invoice("A4"),
                  bg="#E8DFD1", fg=BROWN, relief="flat", font=("Segoe UI", 9), padx=9, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("print_ticket"), command=lambda: self.confirm_and_print_invoice("Ticket"),
                  bg="#E8DFD1", fg=BROWN, relief="flat", font=("Segoe UI", 9), padx=9, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("print_sales_list"), command=self.confirm_and_print_sales_report, bg="#E8DFD1", fg=BROWN,
                  relief="flat", font=("Segoe UI", 9), padx=9, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("export_excel"), command=self.export_sales_excel, bg="#217A4B", fg="white",
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=9, pady=6, cursor="hand2").pack(side=side, padx=3)

        search_frame = tk.Frame(bar, bg=BG)
        search_frame.pack(side="right" if CURRENT_LANG=="fr" else "left", padx=5)
        tk.Label(search_frame, text=f"🔍 {t('search')}:", bg=BG, font=("Segoe UI", 9, "bold")).pack(side="left", padx=3)
        self.sales_search = tk.Entry(search_frame, font=("Segoe UI", 9), width=16)
        self.sales_search.pack(side="left", padx=3, ipady=3)
        self.sales_search.bind("<KeyRelease>", lambda e: self.filter_sales())

        cols = ("id", "date", "customer", "total", "paid", "due", "method")
        heads = (t("invoice_num"), t("date"), t("customer"), t("total_net"), t("paid"), t("due"), t("payment_method"))
        self.sales_tree = self.create_tree(self.body, cols, heads, [65, 110, 230, 130, 130, 130, 110], 9)
        self.sales_tree.bind("<Double-1>", lambda e: self.confirm_and_print_invoice(self.db.get_setting("invoice_format", "A4")))

        summary_bar = tk.Frame(self.body, bg="#EDE4D6", height=32, highlightbackground=BORDER, highlightthickness=1)
        summary_bar.pack(fill="x", pady=(6, 0))
        self.sales_summary_label = tk.Label(summary_bar, text="", bg="#EDE4D6", fg=BROWN, font=("Segoe UI", 9, "bold"))
        self.sales_summary_label.pack(side="left" if CURRENT_LANG=="fr" else "right", padx=15, pady=4)
        self.refresh_sales()

    def filter_sales(self):
        q = self.sales_search.get() if hasattr(self, "sales_search") else ""
        self.refresh_sales(q)

    def refresh_sales(self, query=""):
        if not hasattr(self, "sales_tree") or not self.sales_tree.winfo_exists(): return
        self.sales_tree.delete(*self.sales_tree.get_children())
        if query:
            q = f"%{query.strip()}%"
            rows = self.db.all("""SELECT s.*, COALESCE(c.name, ?) customer FROM sales s
                                  LEFT JOIN customers c ON c.id=s.customer_id
                                  WHERE CAST(s.id AS TEXT) LIKE ? OR c.name LIKE ? OR s.sale_date LIKE ? OR s.payment_method LIKE ?
                                  ORDER BY s.id ASC""", (t("counter_customer"), q, q, q, q))
        else:
            rows = self.db.all("""SELECT s.*, COALESCE(c.name, ?) customer FROM sales s
                                  LEFT JOIN customers c ON c.id=s.customer_id ORDER BY s.id ASC""", (t("counter_customer"),))
        for r in rows:
            self.sales_tree.insert("", "end", values=(
                r["id"], r["sale_date"], r["customer"],
                money(r["total"]), money(r["paid"]), money(r["total"] - r["paid"]),
                r["payment_method"]
            ))
        if hasattr(self, "sales_summary_label") and self.sales_summary_label.winfo_exists():
            tot_sum = sum(r["total"] for r in rows)
            tot_paid = sum(r["paid"] for r in rows)
            tot_due = sum(r["total"] - r["paid"] for r in rows)
            lbl = f"📊 إجمالي الفواتير: {len(rows)} | المجموع: {money(tot_sum)} | المدفوع: {money(tot_paid)} | الديون: {money(tot_due)}" if CURRENT_LANG=="ar" else f"📊 Factures : {len(rows)} | Total : {money(tot_sum)} | Payé : {money(tot_paid)} | Crédits : {money(tot_due)}"
            self.sales_summary_label.config(text=lbl)

    def selected_sale_ids(self):
        sel = self.sales_tree.selection()
        if not sel:
            messagebox.showwarning(t("app_title"), "Sélectionnez au moins une facture" if CURRENT_LANG=="fr" else "يرجى تحديد فاتورة أو أكثر من الجدول أولاً !")
            return []
        ids = []
        for item in sel:
            try:
                ids.append(int(self.sales_tree.item(item, "values")[0]))
            except Exception:
                pass
        return ids

    def selected_sale_id(self):
        ids = self.selected_sale_ids()
        return ids[0] if ids else None

    def edit_sale_dialog(self):
        sid = self.selected_sale_id()
        if not sid: return
        s = self.db.one("SELECT * FROM sales WHERE id=?", (sid,))
        win = tk.Toplevel(self)
        win.title(f"{t('edit_sale')} N° {sid}")
        center_window(win, 460, 290)
        win.configure(bg=BG)

        bottom_bar = tk.Frame(win, bg="#EDE4D6", height=48, highlightbackground=BORDER, highlightthickness=1)
        bottom_bar.pack(side="bottom", fill="x")
        bottom_bar.pack_propagate(False)

        tk.Label(win, text=f"Facture N° {sid} - {t('total_net')}: {money(s['total'])}", bg=BG, fg=BROWN, font=("Segoe UI", 11, "bold")).pack(pady=8)
        f = tk.Frame(win, bg=BG); f.pack(padx=15, pady=2)
        tk.Label(f, text=t("date"), bg=BG, font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=6, pady=4, sticky="e")
        dt_entry = tk.Entry(f, width=18, font=("Segoe UI", 9)); dt_entry.grid(row=0, column=1, padx=6, pady=4)
        dt_entry.insert(0, s["sale_date"])
        tk.Label(f, text=t("paid"), bg=BG, font=("Segoe UI", 9, "bold")).grid(row=1, column=0, padx=6, pady=4, sticky="e")
        paid_entry = tk.Entry(f, width=18, font=("Segoe UI", 9)); paid_entry.grid(row=1, column=1, padx=6, pady=4)
        paid_entry.insert(0, str(s["paid"]))
        tk.Label(f, text=t("payment_method"), bg=BG, font=("Segoe UI", 9, "bold")).grid(row=2, column=0, padx=6, pady=4, sticky="e")
        meth_combo = ttk.Combobox(f, values=["Espèces", "Virement", "Chèque", "Crédit"], state="readonly", width=16)
        meth_combo.set(s["payment_method"])
        meth_combo.grid(row=2, column=1, padx=6, pady=4)
        tk.Label(f, text=t("notes"), bg=BG, font=("Segoe UI", 9, "bold")).grid(row=3, column=0, padx=6, pady=4, sticky="e")
        notes_entry = tk.Entry(f, width=22, font=("Segoe UI", 9)); notes_entry.grid(row=3, column=1, padx=6, pady=4)
        notes_entry.insert(0, s["notes"] or "")

        def save():
            pd = num(paid_entry.get())
            if pd < 0 or pd > s["total"]:
                beep_err(); messagebox.showerror(t("app_title"), "Montant payé invalide" if CURRENT_LANG=="fr" else "المبلغ المدفوع غير صالح", parent=win); return
            msg = f"هل أنت متأكد من حفظ التعديلات على الفاتورة رقم {sid}؟" if CURRENT_LANG=="ar" else f"Confirmer les modifications de la facture N° {sid} ?"
            if not messagebox.askyesno(t("app_title"), msg, parent=win):
                return
            self.db.run("""UPDATE sales SET sale_date=?, paid=?, payment_method=?, notes=? WHERE id=?""",
                        (dt_entry.get().strip(), pd, meth_combo.get(), notes_entry.get().strip(), sid))
            beep_ok()
            win.destroy()
            self.refresh_sales()
            messagebox.showinfo(t("app_title"), t("saved_success"))

        def cancel():
            if messagebox.askyesno(t("app_title"), "هل أنت متأكد من إلغاء التعديل والخروج؟" if CURRENT_LANG=="ar" else "Annuler les modifications ?", parent=win):
                win.destroy()

        btn_box = tk.Frame(bottom_bar, bg="#EDE4D6")
        btn_box.pack(expand=True)
        tk.Button(btn_box, text="💾 حفظ التعديلات" if CURRENT_LANG=="ar" else "💾 Enregistrer",
                  command=save, bg=GOLD, fg=BROWN, relief="flat", font=("Segoe UI", 9, "bold"), padx=14, pady=4, cursor="hand2").pack(side="left", padx=6, pady=8)
        tk.Button(btn_box, text="❌ إلغاء" if CURRENT_LANG=="ar" else "❌ Annuler",
                  command=cancel, bg="#E8DFD1", fg=BROWN, relief="flat", font=("Segoe UI", 9), padx=14, pady=4, cursor="hand2").pack(side="left", padx=6, pady=8)

    def delete_sale_action(self):
        sids = self.selected_sale_ids()
        if not sids: return
        count = len(sids)
        if count == 1:
            msg = f"هل أنت متأكد من حذف الفاتورة رقم {sids[0]} نهائياً؟\n\n• سيتم استرجاع كميات السلع المباعة إلى المخزون تلقائياً." if CURRENT_LANG=="ar" else f"Supprimer définitivement la facture N° {sids[0]} ?\n\n• Le stock sera restauré."
        else:
            msg = f"هل أنت متأكد من حذف {count} فواتير محددة نهائياً؟\n\n• سيتم استرجاع جميع كميات السلع المباعة إلى المخزون." if CURRENT_LANG=="ar" else f"Supprimer définitivement les {count} factures sélectionnées ?\n\n• Le stock sera restauré."
        if not messagebox.askyesno(t("app_title"), msg):
            return
        for sid in sids:
            items = self.db.all("SELECT product_id, quantity FROM sale_items WHERE sale_id=?", (sid,))
            for it in items:
                if it["product_id"]:
                    self.db.run("UPDATE products SET stock = stock + ? WHERE id = ?", (it["quantity"], it["product_id"]))
            self.db.run("DELETE FROM sale_items WHERE sale_id=?", (sid,))
            self.db.run("DELETE FROM sales WHERE id=?", (sid,))
        self.db.resequence_table("sales")
        beep_ok()
        self.refresh_sales()
        messagebox.showinfo(t("app_title"), f"تم حذف {count} فاتورة بنجاح !" if CURRENT_LANG=="ar" else f"{count} facture(s) supprimée(s) avec succès !")

    def confirm_and_print_invoice(self, format_type):
        sid = self.selected_sale_id()
        if not sid: return
        fmt_name = "A4" if format_type == "A4" else ("وصل حراري 80mm" if CURRENT_LANG=="ar" else "Ticket 80mm")
        msg = f"هل أنت متأكد من رغبتك في طباعة الفاتورة رقم {sid} بصيغة ({fmt_name})؟" if CURRENT_LANG=="ar" else f"Imprimer la facture N° {sid} au format ({fmt_name}) ?"
        if messagebox.askyesno(t("app_title"), msg):
            self.generate_invoice_html(sid, format_type=format_type, open_browser=True)

    def confirm_and_print_sales_report(self):
        msg = "هل أنت متأكد من رغبتك في طباعة واستخراج تقرير المبيعات الكامل؟" if CURRENT_LANG=="ar" else "Imprimer le rapport complet des ventes ?"
        if messagebox.askyesno(t("app_title"), msg):
            self.print_sales_report()

    def print_sales_report(self):
        rows = self.db.all("""SELECT s.id, s.sale_date, COALESCE(c.name, ?) customer, s.total, s.paid, (s.total-s.paid) due, s.payment_method
                              FROM sales s LEFT JOIN customers c ON c.id=s.customer_id ORDER BY s.id ASC""", (t("counter_customer"),))
        data = [[r["id"], r["sale_date"], r["customer"], money(r["total"]), money(r["paid"]), money(r["due"]), r["payment_method"]] for r in rows]
        headers = ["N°", t("date"), t("customer"), t("total_net"), t("paid"), t("due"), t("payment_method")]
        self.print_table_report(t("print_sales_list"), headers, data, "journal_ventes")

    # --- فاتورة بيع جديدة ---
    def new_sale_modal(self):
        products = self.db.all("SELECT * FROM products WHERE stock>0 ORDER BY name")
        if not products:
            messagebox.showwarning(t("app_title"), "لا توجد منتجات متوفرة بالمخزون حالياً!" if CURRENT_LANG=="ar" else "Aucun produit en stock!")
            return
        win = tk.Toplevel(self)
        win.title(t("new_invoice"))
        center_window(win, 960, 580)
        win.configure(bg=BG)

        total_net_var = tk.StringVar(value=money(0))
        total_qty_var = tk.StringVar(value="0")

        bottom = tk.Frame(win, bg="#EDE4D6", padx=12, pady=6, highlightbackground=BORDER, highlightthickness=1)
        bottom.pack(side="bottom", fill="x")

        screen_frame = tk.Frame(win, bg="#1E0E08", highlightbackground=GOLD, highlightthickness=2, padx=12, pady=6)
        screen_frame.pack(fill="x", padx=12, pady=(8, 4))
        top_info = tk.Frame(screen_frame, bg="#1E0E08")
        top_info.pack(fill="x")

        client_box = tk.Frame(top_info, bg="#1E0E08")
        client_box.pack(side="left" if CURRENT_LANG=="fr" else "right")
        tk.Label(client_box, text=f"👤 {t('customer')}:", bg="#1E0E08", fg=GOLD, font=("Segoe UI", 9, "bold")).pack(side="left", padx=4)
        custs = self.db.all("SELECT * FROM customers ORDER BY name")
        cust_choices = [t("counter_customer")] + [f"{c['id']} - {c['name']}" for c in custs]
        cust_combo = ttk.Combobox(client_box, values=cust_choices, state="readonly", width=22)
        cust_combo.current(0)
        cust_combo.pack(side="left", padx=4)
        tk.Label(client_box, text=f"📅 {t('date')}:", bg="#1E0E08", fg=GOLD, font=("Segoe UI", 9, "bold")).pack(side="left", padx=4)
        dt_entry = tk.Entry(client_box, width=11, justify="center", font=("Segoe UI", 9, "bold"))
        dt_entry.insert(0, today())
        dt_entry.pack(side="left", padx=4)

        disp_box = tk.Frame(top_info, bg="#1E0E08")
        disp_box.pack(side="right" if CURRENT_LANG=="fr" else "left")
        net_screen = tk.Frame(disp_box, bg="#2D150B", padx=10, pady=2, highlightbackground="#542B1A", highlightthickness=1)
        net_screen.pack(side="right", padx=4)
        tk.Label(net_screen, text="TOTAL NET / المجموع الصافي", bg="#2D150B", fg=GOLD, font=("Segoe UI", 8, "bold")).pack(anchor="e")
        tk.Label(net_screen, textvariable=total_net_var, bg="#2D150B", fg="#4EF08B", font=("Segoe UI", 15, "bold")).pack(anchor="e")
        prod_screen = tk.Frame(disp_box, bg="#2D150B", padx=10, pady=2, highlightbackground="#542B1A", highlightthickness=1)
        prod_screen.pack(side="right", padx=4)
        tk.Label(prod_screen, text="TOTAL PRODUITS / مجموع المنتجات", bg="#2D150B", fg="#E2D7CF", font=("Segoe UI", 8, "bold")).pack(anchor="e")
        tk.Label(prod_screen, textvariable=total_qty_var, bg="#2D150B", fg="#FFD700", font=("Segoe UI", 15, "bold")).pack(anchor="e")

        entry_bar = tk.LabelFrame(win, text=f"📦 {t('designation')} & {t('charg')} & {t('quantity')}", bg="#FFFBF2", fg=BROWN,
                                  font=("Segoe UI", 9, "bold"), padx=10, pady=4)
        entry_bar.pack(fill="x", padx=12, pady=3)

        tk.Label(entry_bar, text="Code-barres:", bg="#FFFBF2", font=("Segoe UI", 8, "bold")).grid(row=0, column=0, padx=2, sticky="w")
        barcode_entry = tk.Entry(entry_bar, font=("Segoe UI", 10, "bold"), width=13, bg="white", relief="solid", bd=1)
        barcode_entry.grid(row=1, column=0, padx=2, ipady=2)
        barcode_entry.focus_set()

        tk.Label(entry_bar, text="Désignation (المنتج):", bg="#FFFBF2", font=("Segoe UI", 8, "bold")).grid(row=0, column=1, padx=2, sticky="w")
        prod_map = {f"{p['id']} - {p['name']} (Stock: {p['stock']})": p for p in products}
        prod_combo = ttk.Combobox(entry_bar, values=list(prod_map), state="readonly", width=26)
        prod_combo.current(0)
        prod_combo.grid(row=1, column=1, padx=2)

        tk.Label(entry_bar, text="Charg. (الحمولة):", bg="#FFFBF2", font=("Segoe UI", 8, "bold")).grid(row=0, column=2, padx=2, sticky="w")
        charg_entry = tk.Entry(entry_bar, width=6, justify="center", font=("Segoe UI", 9, "bold"))
        charg_entry.insert(0, "1")
        charg_entry.grid(row=1, column=2, padx=2, ipady=2)

        tk.Label(entry_bar, text="Total (الكمية):", bg="#FFFBF2", font=("Segoe UI", 8, "bold")).grid(row=0, column=3, padx=2, sticky="w")
        qty_entry = tk.Entry(entry_bar, width=7, justify="center", font=("Segoe UI", 9, "bold"))
        qty_entry.insert(0, "1")
        qty_entry.grid(row=1, column=3, padx=2, ipady=2)

        tk.Label(entry_bar, text="Prix Uni (السعر):", bg="#FFFBF2", font=("Segoe UI", 8, "bold")).grid(row=0, column=4, padx=2, sticky="w")
        price_entry = tk.Entry(entry_bar, width=9, justify="center", font=("Segoe UI", 9, "bold"))
        price_entry.grid(row=1, column=4, padx=2, ipady=2)

        tk.Label(entry_bar, text="Total Produit:", bg="#FFFBF2", font=("Segoe UI", 8, "bold")).grid(row=0, column=5, padx=2, sticky="w")
        total_prod_lbl = tk.Label(entry_bar, text="0.00 DA", bg="#FFFBF2", fg=GREEN, font=("Segoe UI", 10, "bold"), width=12)
        total_prod_lbl.grid(row=1, column=5, padx=2)

        def update_calc(*_):
            q = num(qty_entry.get())
            pr = num(price_entry.get())
            total_prod_lbl.config(text=money(q * pr))

        qty_entry.bind("<KeyRelease>", update_calc)
        price_entry.bind("<KeyRelease>", update_calc)

        def sync_selected_product(_=None):
            p = prod_map.get(prod_combo.get())
            if p:
                price_entry.delete(0, "end"); price_entry.insert(0, str(p["sale_price"]))
                charg_entry.delete(0, "end"); charg_entry.insert(0, str(p.get("charg") or 1))
                update_calc()

        prod_combo.bind("<<ComboboxSelected>>", sync_selected_product)
        sync_selected_product()

        cols = ("num", "name", "charg", "qty", "price", "subtotal")
        heads = ("N°", "Désignation", "Charg.", "Total", "Prix Uni.", "Total Produit")
        cart_tree = self.create_tree(win, cols, heads, [40, 310, 80, 80, 110, 130], 6)
        cart_items = []

        def recalc_screens():
            tot_net = sum(x["subtotal"] for x in cart_items)
            tot_qty = sum(x["qty"] for x in cart_items)
            total_net_var.set(money(tot_net))
            total_qty_var.set(f"{tot_qty:g}")
            paid_entry.delete(0, "end")
            paid_entry.insert(0, str(tot_net))

        def add_item_to_cart(prod, chg, q, pr):
            if not prod or q <= 0 or pr < 0:
                beep_err(); messagebox.showerror(t("app_title"), t("field_required"), parent=win); return False
            already_in_cart = sum(x["qty"] for x in cart_items if x["p"]["id"] == prod["id"])
            if q + already_in_cart > prod["stock"]:
                beep_err(); messagebox.showerror(t("app_title"), f"{t('stock_insufficient')}\n(Stock: {prod['stock']})", parent=win); return False
            found = False
            for it in cart_items:
                if it["p"]["id"] == prod["id"] and it["price"] == pr:
                    it["charg"] += chg
                    it["qty"] += q
                    it["subtotal"] = it["qty"] * it["price"]
                    found = True
                    break
            if not found:
                cart_items.append({"p": prod, "charg": chg, "qty": q, "price": pr, "subtotal": q * pr})
            cart_tree.delete(*cart_tree.get_children())
            for idx, it in enumerate(cart_items, start=1):
                cart_tree.insert("", "end", values=(idx, it["p"]["name"], it["charg"], f"{it['qty']:g}", money(it["price"]), money(it["subtotal"])))
            recalc_screens()
            beep_ok()
            return True

        def on_barcode_enter(e=None):
            code = barcode_entry.get().strip()
            if not code: return
            p_found = self.db.one("SELECT * FROM products WHERE barcode=?", (code,))
            if p_found:
                chg = num(charg_entry.get() or "1")
                q = num(qty_entry.get() or "1")
                add_item_to_cart(p_found, chg, q, p_found["sale_price"])
                barcode_entry.delete(0, "end")
            else:
                beep_err(); messagebox.showwarning(t("app_title"), f"{t('product_not_found')}\n({code})", parent=win)
                barcode_entry.select_range(0, "end")

        barcode_entry.bind("<Return>", on_barcode_enter)

        def add_btn_click():
            p = prod_map.get(prod_combo.get())
            if p:
                add_item_to_cart(p, num(charg_entry.get() or "1"), num(qty_entry.get()), num(price_entry.get()))

        def delete_selected_cart_line():
            sel = cart_tree.selection()
            if not sel:
                messagebox.showwarning(t("app_title"), "Sélectionnez une ligne à supprimer"); return
            idx = int(cart_tree.item(sel[0], "values")[0]) - 1
            if 0 <= idx < len(cart_items):
                cart_items.pop(idx)
                cart_tree.delete(*cart_tree.get_children())
                for i, it in enumerate(cart_items, start=1):
                    cart_tree.insert("", "end", values=(i, it["p"]["name"], it["charg"], f"{it['qty']:g}", money(it["price"]), money(it["subtotal"])))
                recalc_screens()
                beep_ok()

        btn_box = tk.Frame(entry_bar, bg="#FFFBF2")
        btn_box.grid(row=1, column=6, padx=4)
        tk.Button(btn_box, text=f"➕ {t('add_to_cart')}", command=add_btn_click,
                  bg=GOLD, fg=BROWN, font=("Segoe UI", 9, "bold"), relief="flat", padx=8, pady=2, cursor="hand2").pack(side="left", padx=2)
        tk.Button(btn_box, text=f"🗑️ {t('delete_line')}", command=delete_selected_cart_line,
                  bg="#F7D8D5", fg=RED, font=("Segoe UI", 9, "bold"), relief="flat", padx=6, pady=2, cursor="hand2").pack(side="left", padx=2)

        pay_row = tk.Frame(bottom, bg="#EDE4D6")
        pay_row.pack(fill="x", pady=2)
        tk.Label(pay_row, text=f"{t('paid')}:", bg="#EDE4D6", font=("Segoe UI", 9, "bold")).pack(side="left", padx=4)
        paid_entry = tk.Entry(pay_row, width=12, justify="center", font=("Segoe UI", 10, "bold"))
        paid_entry.insert(0, "0")
        paid_entry.pack(side="left", padx=4)

        def set_pay_full():
            tot = sum(x["subtotal"] for x in cart_items)
            paid_entry.delete(0, "end"); paid_entry.insert(0, str(tot))
        def set_pay_zero():
            paid_entry.delete(0, "end"); paid_entry.insert(0, "0")

        tk.Button(pay_row, text=t("pay_full"), command=set_pay_full, bg="#DFD4C2", font=("Segoe UI", 8), relief="flat", padx=5).pack(side="left", padx=2)
        tk.Button(pay_row, text=t("pay_zero"), command=set_pay_zero, bg="#F7D8D5", font=("Segoe UI", 8), relief="flat", padx=5).pack(side="left", padx=2)
        tk.Label(pay_row, text=f"{t('payment_method')}:", bg="#EDE4D6", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(12, 4))
        method_combo = ttk.Combobox(pay_row, values=["Espèces", "Virement", "Chèque", "Crédit"], state="readonly", width=11)
        method_combo.set("Espèces"); method_combo.pack(side="left", padx=4)
        tk.Label(pay_row, text=f"{t('notes')}:", bg="#EDE4D6", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(12, 4))
        notes_entry = tk.Entry(pay_row, width=22, font=("Segoe UI", 9))
        notes_entry.pack(side="left", padx=4)

        def save_sale():
            if not cart_items:
                beep_err(); messagebox.showwarning(t("app_title"), "يرجى إضافة منتج واحد على الأقل للسلة" if CURRENT_LANG=="ar" else "Ajoutez au moins un produit", parent=win); return
            tot = sum(x["subtotal"] for x in cart_items)
            pd = num(paid_entry.get())
            if pd < 0 or pd > tot:
                beep_err(); messagebox.showerror(t("app_title"), "Montant payé invalide" if CURRENT_LANG=="fr" else "المبلغ المدفوع غير صالح", parent=win); return
            msg = (f"هل أنت متأكد من تأكيد وتسجيل عملية البيع؟\n\n• المجموع الصافي: {money(tot)}\n• المبلغ المدفوع: {money(pd)}\n• المتبقي (دين): {money(tot - pd)}") if CURRENT_LANG=="ar" else (f"Confirmer et enregistrer cette vente ?\n\n• Total Net : {money(tot)}\n• Montant Payé : {money(pd)}\n• Reste (Dû) : {money(tot - pd)}")
            if not messagebox.askyesno(t("app_title"), msg, parent=win):
                return
            cid = None
            if cust_combo.get() != t("counter_customer"):
                cid = int(cust_combo.get().split(" - ")[0])
            try:
                cur = self.db.conn.cursor()
                cur.execute("""INSERT INTO sales(sale_date, customer_id, total, paid, payment_method, notes)
                               VALUES(?, ?, ?, ?, ?, ?)""",
                            (dt_entry.get().strip() or today(), cid, tot, pd, method_combo.get(), notes_entry.get()))
                sid = cur.lastrowid
                for it in cart_items:
                    p = it["p"]
                    cur.execute("""INSERT INTO sale_items(sale_id, product_id, product_name, charg, quantity, unit_price, unit_cost, subtotal)
                                   VALUES(?, ?, ?, ?, ?, ?, ?, ?)""",
                                (sid, p["id"], p["name"], it["charg"], it["qty"], it["price"], p["purchase_price"], it["subtotal"]))
                    cur.execute("UPDATE products SET stock=stock-? WHERE id=?", (it["qty"], p["id"]))
                self.db.conn.commit()
            except Exception as e:
                self.db.conn.rollback()
                messagebox.showerror(t("app_title"), str(e), parent=win)
                return
            beep_ok()
            win.destroy()
            self.refresh_sales()
            pref_fmt = self.db.get_setting("invoice_format", "A4")
            self.generate_invoice_html(sid, format_type=pref_fmt, open_browser=True)

        def cancel_sale():
            if cart_items:
                msg = "توجد سلع في السلة الحالية، هل أنت متأكد من إلغاء عملية البيع والخروج؟" if CURRENT_LANG=="ar" else "Il y a des articles dans le panier, voulez-vous vraiment annuler la vente ?"
                if not messagebox.askyesno(t("app_title"), msg, parent=win):
                    return
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", cancel_sale)

        btn_box = tk.Frame(bottom, bg="#EDE4D6")
        btn_box.pack(expand=True)
        tk.Button(btn_box, text=f"💾 {t('save_sale')}", command=save_sale, bg=GOLD, fg=BROWN,
                  font=("Segoe UI", 10, "bold"), relief="flat", padx=18, pady=5, cursor="hand2").pack(side="left", padx=8, pady=4)
        tk.Button(btn_box, text="❌ إلغاء البيع" if CURRENT_LANG=="ar" else "❌ Annuler la vente",
                  command=cancel_sale, bg="#E8DFD1", fg=BROWN,
                  font=("Segoe UI", 9, "bold"), relief="flat", padx=14, pady=5, cursor="hand2").pack(side="left", padx=8, pady=4)

    def generate_invoice_html(self, sid, format_type="A4", open_browser=False):
        s = self.db.one("""SELECT s.*, COALESCE(c.name, ?) customer, COALESCE(c.phone, '') phone, COALESCE(c.address, '') address
                           FROM sales s LEFT JOIN customers c ON c.id=s.customer_id WHERE s.id=?""", (t("counter_customer"), sid))
        if not s: return None
        items = self.db.all("SELECT * FROM sale_items WHERE sale_id=?", (sid,))
        comp_name = html.escape(self.db.get_setting("company_name", APP_NAME))
        comp_act = html.escape(self.db.get_setting("company_activity", t("subtitle")))
        comp_phone = html.escape(self.db.get_setting("company_phone", "0561 88 45 06"))
        comp_addr = html.escape(self.db.get_setting("company_address", ""))
        comp_rc = html.escape(self.db.get_setting("company_rc", ""))
        comp_nif = html.escape(self.db.get_setting("company_nif", ""))
        footer_msg = html.escape(self.db.get_setting("invoice_footer", ""))
        total_qty_sum = sum(i["quantity"] for i in items)

        if format_type == "Ticket":
            rows_html = "".join(f"""
            <tr>
                <td style="text-align:left;padding:4px 0;">{idx+1}. {html.escape(i['product_name'])}<br><small>Charg: {i['charg']} | {i['quantity']:g} x {money(i['unit_price'])}</small></td>
                <td style="text-align:right;vertical-align:bottom;padding:4px 0;"><b>{money(i['subtotal'])}</b></td>
            </tr>
            """ for idx, i in enumerate(items))
            doc = f"""<!doctype html>
            <html lang="fr" dir="ltr">
            <head><meta charset="utf-8"><title>Ticket_{sid}</title>
            <style>
                body {{ font-family: 'Courier New', monospace, sans-serif; width: 76mm; margin: 0 auto; padding: 10px; color: #000; font-size: 13px; }}
                .center {{ text-align: center; }}
                .bold {{ font-weight: bold; }}
                .line {{ border-bottom: 1px dashed #000; margin: 8px 0; }}
                table {{ width: 100%; border-collapse: collapse; }}
                @media print {{ button {{ display: none; }} body {{ width: 76mm; padding: 0; }} }}
            </style></head>
            <body>
                <div class="center">
                    <h2 style="margin:2px 0;">{comp_name}</h2>
                    <div>{comp_act}</div>
                    <div>Tél: {comp_phone}</div>
                    <div>{comp_addr}</div>
                    <div class="line"></div>
                    <div class="bold">TICKET DE CAISSE N° {sid}</div>
                    <div>Date: {s['sale_date']}</div>
                    <div>Client: {html.escape(s['customer'])}</div>
                </div>
                <div class="line"></div>
                <table>{rows_html}</table>
                <div class="line"></div>
                <table>
                    <tr><td>TOTAL PRODUITS:</td><td style="text-align:right;font-weight:bold;">{total_qty_sum:g}</td></tr>
                    <tr><td class="bold">TOTAL NET:</td><td class="bold" style="text-align:right;font-size:15px;">{money(s['total'])}</td></tr>
                    <tr><td>Payé:</td><td style="text-align:right;">{money(s['paid'])}</td></tr>
                    <tr><td class="bold">Reste:</td><td class="bold" style="text-align:right;">{money(s['total'] - s['paid'])}</td></tr>
                </table>
                <div class="line"></div>
                <div class="center" style="font-size:11px;margin-top:10px;">
                    {footer_msg}<br><br>
                    <button onclick="window.print()" style="padding:6px 12px;font-size:12px;cursor:pointer;">Imprimer Ticket</button>
                </div>
            </body></html>"""
        else:
            rows_html = "".join(f"""
            <tr>
                <td style="text-align:center;">{idx+1}</td>
                <td><b>{html.escape(i['product_name'])}</b></td>
                <td style="text-align:center;">{i['charg']}</td>
                <td style="text-align:center;"><b>{i['quantity']:g}</b></td>
                <td style="text-align:right;">{money(i['unit_price'])}</td>
                <td style="text-align:right;"><b>{money(i['subtotal'])}</b></td>
            </tr>
            """ for idx, i in enumerate(items))
            doc = f"""<!doctype html>
            <html lang="fr" dir="ltr">
            <head><meta charset="utf-8"><title>Facture_{sid}</title>
            <style>
                body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 30px; color: #2B160F; }}
                .header-table {{ width: 100%; border-bottom: 3px solid #E09C00; padding-bottom: 15px; }}
                .company-info h1 {{ margin: 0; color: #3B1B12; font-size: 26px; }}
                .company-info p {{ margin: 3px 0; color: #555; font-size: 13px; }}
                .client-box {{ background: #F8F5EE; border: 1px solid #E0D7CB; padding: 12px 18px; border-radius: 6px; margin: 20px 0; }}
                table.items {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
                table.items th {{ background: #3B1B12; color: #FFF; padding: 10px; font-size: 13px; }}
                table.items td {{ border: 1px solid #E0D7CB; padding: 9px; font-size: 13px; }}
                .sum-table {{ width: 380px; margin-left: auto; margin-top: 20px; border-collapse: collapse; }}
                .sum-table td, .sum-table th {{ padding: 8px 12px; border: 1px solid #E0D7CB; }}
                .sum-table th {{ background: #F8F5EE; text-align: left; font-size: 13px; }}
                .footer {{ text-align: center; margin-top: 40px; font-size: 12px; color: #777; border-top: 1px solid #DDD; padding-top: 15px; }}
                @media print {{ button {{ display: none; }} body {{ margin: 15mm; }} }}
            </style></head>
            <body>
                <table class="header-table">
                    <tr>
                        <td class="company-info" style="vertical-align:top;">
                            <h1>{comp_name}</h1>
                            <p><b>{comp_act}</b></p>
                            <p>Tél: {comp_phone} | {comp_addr}</p>
                            {f'<p>RC: {comp_rc} | NIF: {comp_nif}</p>' if comp_rc or comp_nif else ''}
                        </td>
                        <td style="text-align:right;vertical-align:top;">
                            <h2 style="color:#E09C00;margin:0 0 5px 0;">FACTURE N° {sid}</h2>
                            <p><b>Date:</b> {s['sale_date']}</p>
                            <p><b>Mode:</b> {html.escape(s['payment_method'])}</p>
                        </td>
                    </tr>
                </table>
                <div class="client-box">
                    <table style="width:100%;">
                        <tr>
                            <td><b>Client:</b> {html.escape(s['customer'])}</td>
                            <td><b>Téléphone:</b> {html.escape(s['phone'] or '---')}</td>
                            <td><b>Adresse:</b> {html.escape(s['address'] or '---')}</td>
                        </tr>
                    </table>
                </div>
                <table class="items">
                    <thead>
                        <tr>
                            <th style="width:45px;">N°</th>
                            <th>Désignation</th>
                            <th style="width:80px;">Charg.</th>
                            <th style="width:90px;">Total (Qté)</th>
                            <th style="width:120px;">Prix Uni.</th>
                            <th style="width:140px;">Total Produit</th>
                        </tr>
                    </thead>
                    <tbody>{rows_html}</tbody>
                </table>
                <table class="sum-table">
                    <tr><th>Total Produits (مجموع المنتجات):</th><td style="text-align:right;font-weight:bold;font-size:14px;">{total_qty_sum:g}</td></tr>
                    <tr style="background:#FFF9E6;"><th>TOTAL NET (المجموع الصافي):</th><td style="text-align:right;font-weight:bold;font-size:16px;color:#3B1B12;">{money(s['total'])}</td></tr>
                    <tr><th>Montant Payé (المدفوع):</th><td style="text-align:right;color:#217A4B;font-weight:bold;">{money(s['paid'])}</td></tr>
                    <tr style="background:#FFF2F0;"><th>Reste à Payer (المتبقي):</th><td style="text-align:right;color:#B83226;font-weight:bold;">{money(s['total'] - s['paid'])}</td></tr>
                </table>
                <div class="footer">
                    <p>{footer_msg}</p>
                    <button onclick="window.print()" style="background:#E09C00;color:#3B1B12;border:none;padding:10px 22px;font-weight:bold;font-size:14px;cursor:pointer;border-radius:4px;">🖨️ Imprimer la Facture</button>
                </div>
            </body></html>"""

        path = ROOT / f"facture_{sid}_{format_type}.html"
        path.write_text(doc, encoding="utf-8")
        if open_browser: webbrowser.open(path.as_uri())
        return path

    # ==========================================
    # 4. الزبائن
    # ==========================================
    def show_customers(self):
        self.clear()
        self.heading(t("customers"), t("subtitle"))
        bar = tk.Frame(self.body, bg=BG)
        bar.pack(fill="x", pady=(0, 8))
        side = "left" if CURRENT_LANG=="fr" else "right"
        tk.Button(bar, text=t("add_customer"), command=self.add_customer, bg=GOLD, fg=BROWN,
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=12, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("edit_customer"), command=self.edit_customer_dialog, bg="#E8DFD1", fg=BROWN,
                  relief="flat", font=("Segoe UI", 9), padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("delete_customer"), command=self.delete_customer_action, bg="#F7D8D5", fg=RED,
                  relief="flat", font=("Segoe UI", 9), padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("pay_debt"), command=self.pay_customer_debt, bg="#E8DFD1", fg=BROWN,
                  relief="flat", font=("Segoe UI", 9), padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("account_statement"), command=self.customer_statement, bg="#E8DFD1", fg=BROWN,
                  relief="flat", font=("Segoe UI", 9), padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("print_customers_list"), command=self.print_customers_report, bg="#E8DFD1", fg=BROWN,
                  relief="flat", font=("Segoe UI", 9), padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("import_customers_excel"), command=self.import_customers_excel, bg="#1B6396", fg="white",
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("download_customers_template"), command=self.save_customers_template, bg="#E8DFD1", fg=BROWN,
                  relief="flat", font=("Segoe UI", 9), padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("export_excel"), command=self.export_customers_excel, bg="#217A4B", fg="white",
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)

        cols = ("id", "name", "phone", "address", "debt")
        heads = ("N°", t("customer"), t("phone"), t("address"), t("balance_due"))
        self.customer_tree = self.create_tree(self.body, cols, heads, [50, 230, 140, 240, 140], 9)
        self.customer_tree.bind("<Double-1>", lambda e: self.edit_customer_dialog())
        self.refresh_customers()

    def refresh_customers(self):
        if not hasattr(self, "customer_tree") or not self.customer_tree.winfo_exists(): return
        self.customer_tree.delete(*self.customer_tree.get_children())
        for c in self.db.all("SELECT * FROM customers ORDER BY id ASC"):
            due = self.db.one("SELECT COALESCE(SUM(total-paid),0) n FROM sales WHERE customer_id=?", (c["id"],))["n"]
            tag = "indebt" if due > 0 else "normal"
            self.customer_tree.insert("", "end", values=(
                c["id"], c["name"], c["phone"], c["address"], money(due)
            ), tags=(tag,))
        self.customer_tree.tag_configure("indebt", background="#FFF3F2")

    def selected_customer_id(self):
        sel = self.customer_tree.selection()
        if not sel:
            messagebox.showwarning(t("app_title"), "Sélectionnez un client" if CURRENT_LANG=="fr" else "يرجى تحديد زبون")
            return None
        return int(self.customer_tree.item(sel[0], "values")[0])

    def add_customer(self):
        self.customer_dialog()

    def edit_customer_dialog(self):
        cid = self.selected_customer_id()
        if cid:
            row = self.db.one("SELECT * FROM customers WHERE id=?", (cid,))
            self.customer_dialog(row)

    def customer_dialog(self, row=None):
        win = tk.Toplevel(self)
        win.title(t("edit_customer") if row else t("add_customer"))
        center_window(win, 460, 270)
        win.configure(bg=BG)
        win.resizable(False, False)

        bottom_bar = tk.Frame(win, bg="#EDE4D6", height=44)
        bottom_bar.pack(side="bottom", fill="x")

        labels = [t("customer"), t("phone"), t("address"), t("notes")]
        keys = ["name", "phone", "address", "notes"]
        entries = {}
        form = tk.Frame(win, bg=BG)
        form.pack(fill="both", expand=True, padx=15, pady=8)
        for i, (l, k) in enumerate(zip(labels, keys)):
            tk.Label(form, text=l, bg=BG, fg=BROWN, font=("Segoe UI", 9, "bold")).grid(
                row=i, column=0, padx=8, pady=4, sticky="e" if CURRENT_LANG=="fr" else "w")
            e = tk.Entry(form, width=28, font=("Segoe UI", 9))
            e.grid(row=i, column=1, padx=8, pady=4, ipady=2)
            if row: e.insert(0, str(row[k] or ""))
            entries[k] = e

        def save():
            name = entries["name"].get().strip()
            if not name:
                messagebox.showerror(t("app_title"), t("field_required"), parent=win); return
            if row:
                self.db.run("UPDATE customers SET name=?, phone=?, address=?, notes=? WHERE id=?",
                            (name, entries["phone"].get().strip(), entries["address"].get().strip(), entries["notes"].get().strip(), row["id"]))
            else:
                self.db.run("INSERT INTO customers(name, phone, address, notes) VALUES(?, ?, ?, ?)",
                            (name, entries["phone"].get().strip(), entries["address"].get().strip(), entries["notes"].get().strip()))
            beep_ok()
            win.destroy()
            self.refresh_customers()

        tk.Button(bottom_bar, text="💾 حفظ البيانات" if CURRENT_LANG=="ar" else "💾 Enregistrer", command=save,
                  bg=GOLD, fg=BROWN, relief="flat", font=("Segoe UI", 9, "bold"), padx=18, pady=4).pack(pady=7)

    def delete_customer_action(self):
        cid = self.selected_customer_id()
        if not cid: return
        due = self.db.one("SELECT COALESCE(SUM(total-paid),0) n FROM sales WHERE customer_id=?", (cid,))["n"]
        msg = t("customer_debt_warning") if due > 0 else t("confirm_delete")
        if messagebox.askyesno(t("app_title"), msg):
            self.db.run("DELETE FROM customers WHERE id=?", (cid,))
            self.db.resequence_table("customers")
            beep_ok()
            self.refresh_customers()
            messagebox.showinfo(t("app_title"), t("saved_success"))

    def print_customers_report(self):
        rows = self.db.all("SELECT * FROM customers ORDER BY id ASC")
        data = []
        for c in rows:
            due = self.db.one("SELECT COALESCE(SUM(total-paid),0) n FROM sales WHERE customer_id=?", (c["id"],))["n"]
            data.append([c["id"], c["name"], c["phone"] or "-", c["address"] or "-", money(due)])
        headers = ["N°", t("customer"), t("phone"), t("address"), t("balance_due")]
        self.print_table_report(t("print_customers_list"), headers, data, "liste_clients")

    def pay_customer_debt(self):
        cid = self.selected_customer_id()
        if not cid: return
        c = self.db.one("SELECT * FROM customers WHERE id=?", (cid,))
        due = self.db.one("SELECT COALESCE(SUM(total-paid),0) n FROM sales WHERE customer_id=?", (cid,))["n"]
        if due <= 0:
            messagebox.showinfo(t("app_title"), "لا توجد أي ديون على هذا الزبون." if CURRENT_LANG=="ar" else "Ce client n'a aucune dette en cours.")
            return
        win = tk.Toplevel(self)
        win.title(t("pay_debt"))
        center_window(win, 360, 220)
        win.configure(bg=BG)
        bottom_bar = tk.Frame(win, bg="#EDE4D6", height=42)
        bottom_bar.pack(side="bottom", fill="x")
        tk.Label(win, text=f"Client: {c['name']}", bg=BG, fg=BROWN, font=("Segoe UI", 11, "bold")).pack(padx=15, pady=6)
        tk.Label(win, text=f"{t('balance_due')}: {money(due)}", bg=BG, fg=RED, font=("Segoe UI", 10, "bold")).pack()
        e = tk.Entry(win, justify="center", font=("Segoe UI", 11, "bold"), width=15)
        e.insert(0, str(due))
        e.pack(padx=15, pady=8, ipady=3)

        def save():
            amt = num(e.get())
            if amt <= 0 or amt > due:
                beep_err(); messagebox.showerror(t("app_title"), "Montant invalide", parent=win); return
            rem = amt
            for s in self.db.all("SELECT id, total, paid FROM sales WHERE customer_id=? AND total>paid ORDER BY id", (cid,)):
                part = min(rem, s["total"] - s["paid"])
                self.db.run("UPDATE sales SET paid=paid+? WHERE id=?", (part, s["id"]))
                rem -= part
                if rem <= 0: break
            beep_ok()
            win.destroy()
            self.refresh_customers()
            messagebox.showinfo(t("app_title"), t("saved_success"))

        tk.Button(bottom_bar, text="تأكيد التسديد" if CURRENT_LANG=="ar" else "Confirmer",
                  command=save, bg=GOLD, fg=BROWN, relief="flat", font=("Segoe UI", 9, "bold"), padx=15, pady=4).pack(pady=6)

    def customer_statement(self):
        cid = self.selected_customer_id()
        if not cid: return
        c = self.db.one("SELECT * FROM customers WHERE id=?", (cid,))
        sales = self.db.all("SELECT * FROM sales WHERE customer_id=? ORDER BY id DESC", (cid,))
        comp_name = html.escape(self.db.get_setting("company_name", APP_NAME))
        rows_html = "".join(f"""
        <tr>
            <td>{s['sale_date']}</td>
            <td>Facture N° {s['id']}</td>
            <td style="text-align:right;">{money(s['total'])}</td>
            <td style="text-align:right;color:green;">{money(s['paid'])}</td>
            <td style="text-align:right;color:red;font-weight:bold;">{money(s['total']-s['paid'])}</td>
        </tr>
        """ for s in sales)
        total_purchases = sum(s["total"] for s in sales)
        total_paid = sum(s["paid"] for s in sales)
        total_due = total_purchases - total_paid
        doc = f"""<!doctype html><html lang="fr" dir="ltr"><head><meta charset="utf-8">
        <title>Relevé_{c['name']}</title>
        <style>body{{font-family:'Segoe UI',sans-serif;margin:30px;color:#3B1B12;}}h1{{color:#E09C00;}}table{{width:100%;border-collapse:collapse;margin-top:15px;}}th,td{{border:1px solid #DDD;padding:8px;text-align:left;}}th{{background:#3B1B12;color:white;}}@media print{{button{{display:none;}}}}</style>
        </head><body>
        <h1>{comp_name} - Relevé de Compte Client</h1>
        <p><b>Client:</b> {html.escape(c['name'])} | <b>Téléphone:</b> {html.escape(c['phone'])}</p>
        <table><thead><tr><th>Date</th><th>Opération</th><th style="text-align:right;">Total</th><th style="text-align:right;">Payé</th><th style="text-align:right;">Solde Dû</th></tr></thead>
        <tbody>{rows_html}</tbody></table>
        <h3 style="text-align:right;margin-top:20px;color:red;">Solde Total Restant: {money(total_due)}</h3>
        <div style="text-align:center;margin-top:30px;"><button onclick="window.print()">Imprimer Relevé</button></div>
        </body></html>"""
        p = ROOT / f"releve_client_{cid}.html"
        p.write_text(doc, encoding="utf-8")
        webbrowser.open(p.as_uri())

    # ==========================================
    # 2.5 Export / Import Excel (100% hors ligne)
    # ==========================================
    def _require_openpyxl(self):
        if not OPENPYXL_AVAILABLE:
            beep_err()
            messagebox.showerror(t("app_title"), t("excel_missing"))
            return False
        return True

    def _ask_save_xlsx(self, base_name):
        return filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=f"{base_name}_{today()}.xlsx")

    def _write_xlsx(self, path, sheets):
        """sheets = liste de (nom_feuille, headers, rows, rtl). Écrit un classeur stylé."""
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        wb = Workbook()
        wb.remove(wb.active)
        thin = Side(style="thin", color="D8CDBE")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for sheet_name, headers, rows, rtl in sheets:
            ws = wb.create_sheet(title=sheet_name[:31])
            ws.sheet_view.rightToLeft = bool(rtl)
            hfill = PatternFill("solid", fgColor="3B1B12")
            hfont = Font(bold=True, color="FFC52E", size=11)
            for ci, h in enumerate(headers, 1):
                c = ws.cell(row=1, column=ci, value=h)
                c.fill = hfill; c.font = hfont; c.border = border
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.row_dimensions[1].height = 24
            zfill = PatternFill("solid", fgColor="F6F3EC")
            for ri, row in enumerate(rows, 2):
                for ci, v in enumerate(row, 1):
                    c = ws.cell(row=ri, column=ci, value=v)
                    c.border = border
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        c.number_format = "#,##0.00"
                    if ri % 2 == 0:
                        c.fill = zfill
            for ci, h in enumerate(headers, 1):
                maxlen = len(str(h))
                for row in rows[:400]:
                    if ci - 1 < len(row):
                        maxlen = max(maxlen, len(str(row[ci - 1] if row[ci - 1] is not None else "")))
                ws.column_dimensions[get_column_letter(ci)].width = min(42, max(9, maxlen + 3))
            ws.freeze_panes = "A2"
            if rows:
                ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows)+1}"
        wb.save(path)

    def export_products_excel(self):
        if not self._require_openpyxl(): return
        path = self._ask_save_xlsx("produits_trafik_saveurs")
        if not path: return
        try:
            rows = []
            for r in self.db.all("""SELECT p.*, COALESCE(SUM(si.subtotal),0) ca FROM products p
                                    LEFT JOIN sale_items si ON si.product_id=p.id GROUP BY p.id ORDER BY p.id"""):
                val_stock = (r["stock"] / r["charg"] if r["charg"] else r["stock"]) * r["purchase_price"]
                rows.append([r["id"], r["name"], r["category"], r["barcode"] or "", r["charg"],
                             r["purchase_price"], r["sale_price"], r["stock"], r["unit"],
                             r["min_stock"], r["expiry"] or "", round(val_stock, 2),
                             round((r["sale_price"] - r["purchase_price"]) * r["stock"], 2)])
            heads = ["N°", t("designation"), t("category"), t("barcode"), t("charg"),
                     t("purchase_price"), t("sale_price"), t("stock"), t("unit"),
                     t("alert_threshold"), t("expiry_date"), "Valeur Stock", "Marge Potentielle"]
            self._write_xlsx(path, [("Produits", heads, rows, CURRENT_LANG == "ar")])
            beep_ok()
            messagebox.showinfo(t("app_title"),
                                f"تم تصدير {len(rows)} منتجاً إلى:\n{path}" if CURRENT_LANG == "ar"
                                else f"{len(rows)} produits exportés vers :\n{path}")
        except Exception as e:
            beep_err()
            messagebox.showerror(t("app_title"), f"Erreur Export Excel : {e}")

    def export_sales_excel(self):
        if not self._require_openpyxl(): return
        path = self._ask_save_xlsx("ventes_trafik_saveurs")
        if not path: return
        try:
            sales = self.db.all("""SELECT s.*, COALESCE(c.name, ?) customer FROM sales s
                                   LEFT JOIN customers c ON c.id=s.customer_id ORDER BY s.id""",
                                (t("counter_customer"),))
            srows = [[r["id"], r["sale_date"], r["customer"], r["payment_method"],
                      r["total"], r["paid"], r["total"] - r["paid"], r["notes"] or ""] for r in sales]
            sheads = [t("invoice_num"), t("date"), t("customer"), t("payment_method"),
                      t("total_net"), t("paid"), t("due"), t("notes")]
            items = self.db.all("""SELECT si.sale_id, s.sale_date, si.product_name, si.charg,
                                          si.quantity, si.unit_price, si.unit_cost, si.subtotal
                                   FROM sale_items si JOIN sales s ON s.id=si.sale_id ORDER BY si.id""")
            irows = [[r["sale_id"], r["sale_date"], r["product_name"], r["charg"],
                      r["quantity"], r["unit_price"], r["unit_cost"],
                      round(r["subtotal"] - r["unit_cost"] * r["quantity"], 2), r["subtotal"]] for r in items]
            iheads = [t("invoice_num"), t("date"), t("designation"), t("charg"),
                      t("quantity"), t("sale_price"), t("unit_cost"), "Marge", t("total_net")]
            tot = sum(r["total"] for r in sales); paid = sum(r["paid"] for r in sales)
            sum_rows = [["Total Factures", len(sales)], ["Total Ventes", round(tot, 2)],
                        ["Total Payé", round(paid, 2)], ["Total Crédits", round(tot - paid, 2)]]
            self._write_xlsx(path, [
                ("Ventes", sheads, srows, CURRENT_LANG == "ar"),
                ("Lignes Ventes", iheads, irows, CURRENT_LANG == "ar"),
                ("Résumé", ["Indicateur", "Valeur"], sum_rows, False)])
            beep_ok()
            messagebox.showinfo(t("app_title"),
                                f"تم تصدير {len(sales)} فاتورة بنجاح !\n{path}" if CURRENT_LANG == "ar"
                                else f"{len(sales)} factures exportées avec succès !\n{path}")
        except Exception as e:
            beep_err()
            messagebox.showerror(t("app_title"), f"Erreur Export Excel : {e}")

    def export_purchases_excel(self):
        if not self._require_openpyxl(): return
        path = self._ask_save_xlsx("achats_trafik_saveurs")
        if not path: return
        try:
            purch = self.db.all("SELECT * FROM purchases ORDER BY id")
            prows = [[r["id"], r["purchase_date"], r["supplier"] or "", r["total"],
                      r["paid"], r["total"] - r["paid"], r["notes"] or ""] for r in purch]
            pheads = ["N°", t("date"), t("supplier"), t("total_net"), t("paid"), t("due"), t("notes")]
            items = self.db.all("""SELECT pi.purchase_id, p.purchase_date, pi.product_name,
                                          pi.quantity, pi.unit_cost, pi.subtotal
                                   FROM purchase_items pi JOIN purchases p ON p.id=pi.purchase_id ORDER BY pi.id""")
            irows = [[r["purchase_id"], r["purchase_date"], r["product_name"],
                      r["quantity"], r["unit_cost"], r["subtotal"]] for r in items]
            iheads = ["N°", t("date"), t("designation"), t("quantity"), t("unit_cost"), t("total_net")]
            self._write_xlsx(path, [
                ("Achats", pheads, prows, CURRENT_LANG == "ar"),
                ("Lignes Achats", iheads, irows, CURRENT_LANG == "ar")])
            beep_ok()
            messagebox.showinfo(t("app_title"),
                                f"تم تصدير {len(purch)} عملية شراء !\n{path}" if CURRENT_LANG == "ar"
                                else f"{len(purch)} achats exportés !\n{path}")
        except Exception as e:
            beep_err()
            messagebox.showerror(t("app_title"), f"Erreur Export Excel : {e}")

    def export_customers_excel(self):
        if not self._require_openpyxl(): return
        path = self._ask_save_xlsx("clients_trafik_saveurs")
        if not path: return
        try:
            rows = []
            for c in self.db.all("SELECT * FROM customers ORDER BY id"):
                due = self.db.one("SELECT COALESCE(SUM(total-paid),0) n FROM sales WHERE customer_id=?",
                                  (c["id"],))["n"]
                nb = self.db.one("SELECT COUNT(*) n FROM sales WHERE customer_id=?", (c["id"],))["n"]
                rows.append([c["id"], c["name"], c["phone"], c["address"], nb, round(due, 2), c["notes"] or ""])
            heads = ["N°", t("customer"), t("phone"), t("address"), "Nb Factures", t("balance_due"), t("notes")]
            self._write_xlsx(path, [("Clients", heads, rows, CURRENT_LANG == "ar")])
            beep_ok()
            messagebox.showinfo(t("app_title"),
                                f"تم تصدير {len(rows)} زبوناً !\n{path}" if CURRENT_LANG == "ar"
                                else f"{len(rows)} clients exportés !\n{path}")
        except Exception as e:
            beep_err()
            messagebox.showerror(t("app_title"), f"Erreur Export Excel : {e}")

    def save_customers_template(self):
        if not self._require_openpyxl(): return
        path = self._ask_save_xlsx("modele_import_clients")
        if not path: return
        try:
            heads = ["الاسم / Nom", "الهاتف / Téléphone", "العنوان / Adresse", "ملاحظات / Notes"]
            demo = [["Client Exemple", "0561000000", "Alger", "—"]]
            self._write_xlsx(path, [("Clients", heads, demo, False)])
            beep_ok()
            messagebox.showinfo(t("app_title"),
                                "تم حفظ نموذج استيراد الزبائن !" if CURRENT_LANG == "ar"
                                else "Le modèle d'importation des clients a été enregistré !")
        except Exception as e:
            messagebox.showerror(t("app_title"), f"Erreur : {e}")

    def _read_tabular_file(self, filepath):
        """Retourne les lignes (liste de tuples) d'un fichier xlsx/xls/csv — hors ligne."""
        if filepath.lower().endswith(".csv"):
            with open(filepath, "r", encoding="utf-8-sig", newline="") as f:
                sample = f.read(2048); f.seek(0)
                delim = ";" if sample.count(";") >= sample.count(",") else ","
                return list(csv.reader(f, delimiter=delim))
        if not OPENPYXL_AVAILABLE:
            raise RuntimeError(t("excel_missing"))
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active
        return [tuple("" if c is None else c for c in row) for row in ws.iter_rows(values_only=True)]

    def import_customers_excel(self):
        filepath = filedialog.askopenfilename(
            title="Sélectionner le fichier Clients / اختر ملف إكسل للزبائن",
            filetypes=[("Excel & CSV", "*.xlsx *.xls *.csv"), ("Tous les fichiers", "*.*")])
        if not filepath: return
        try:
            rows = self._read_tabular_file(filepath)
            if not rows or len(rows) < 2:
                messagebox.showerror(t("app_title"), "Le fichier est vide."); return
            header = [str(c or "").strip().lower() for c in rows[0]]
            col = {}
            for idx, h in enumerate(header):
                if any(k in h for k in ["nom", "client", "اسم", "زبون", "customer"]):
                    col.setdefault("name", idx)
                elif any(k in h for k in ["tel", "phone", "هاتف", "جوال", "موبايل"]):
                    col.setdefault("phone", idx)
                elif any(k in h for k in ["adresse", "عنوان", "address"]):
                    col.setdefault("address", idx)
                elif any(k in h for k in ["note", "remarque", "ملاحظ"]):
                    col.setdefault("notes", idx)
            if "name" not in col:
                messagebox.showerror(t("app_title"),
                                     "Colonne 'Nom / الاسم' introuvable dans le fichier."
                                     if CURRENT_LANG == "fr" else "عمود 'الاسم' غير موجود في الملف.")
                return
            added, updated, skipped = 0, 0, 0
            for r in rows[1:]:
                if not r or all(str(x or "").strip() == "" for x in r): continue
                name = str(r[col["name"]] or "").strip()
                if not name: continue
                phone = str(r[col["phone"]] or "").strip() if "phone" in col and col["phone"] < len(r) else ""
                addr = str(r[col["address"]] or "").strip() if "address" in col and col["address"] < len(r) else ""
                notes = str(r[col["notes"]] or "").strip() if "notes" in col and col["notes"] < len(r) else ""
                exist = self.db.one("SELECT id FROM customers WHERE name=?", (name,))
                if exist:
                    self.db.run("UPDATE customers SET phone=?, address=?, notes=? WHERE id=?",
                                (phone, addr, notes, exist["id"]))
                    updated += 1
                else:
                    self.db.run("INSERT INTO customers(name, phone, address, notes) VALUES(?,?,?,?)",
                                (name, phone, addr, notes))
                    added += 1
            beep_ok()
            msg = (f"✅ زبائن جدد: {added}\n🔄 تم تحديثهم: {updated}" if CURRENT_LANG == "ar"
                   else f"✅ Nouveaux clients : {added}\n🔄 Mis à jour : {updated}")
            messagebox.showinfo(t("app_title"), msg)
            if hasattr(self, "refresh_customers"):
                self.refresh_customers()
        except Exception as e:
            beep_err()
            messagebox.showerror(t("app_title"), f"Erreur lors de l'import : {e}")

    def export_report_excel(self, a, b):
        if not self._require_openpyxl(): return
        path = self._ask_save_xlsx(f"rapport_{a}_{b}")
        if not path: return
        try:
            mvt = []
            for r in self.db.all("SELECT id, sale_date, total, paid FROM sales WHERE sale_date BETWEEN ? AND ?", (a, b)):
                mvt.append([r["sale_date"], "Vente", r["id"], r["total"], r["paid"], r["total"] - r["paid"]])
            for r in self.db.all("SELECT id, purchase_date, total, paid FROM purchases WHERE purchase_date BETWEEN ? AND ?", (a, b)):
                mvt.append([r["purchase_date"], "Achat", r["id"], r["total"], r["paid"], r["total"] - r["paid"]])
            for r in self.db.all("SELECT expense_date, title, amount FROM expenses WHERE expense_date BETWEEN ? AND ?", (a, b)):
                mvt.append([r["expense_date"], "Dépense", r["title"], r["amount"], r["amount"], 0])
            mheads = [t("date"), "Type", "Réf", t("total_net"), t("paid"), t("due")]
            ts = self.db.one("SELECT COALESCE(SUM(total),0) tt, COALESCE(SUM(paid),0) pp FROM sales WHERE sale_date BETWEEN ? AND ?", (a, b))
            tp = self.db.one("SELECT COALESCE(SUM(total),0) tt FROM purchases WHERE purchase_date BETWEEN ? AND ?", (a, b))
            te = self.db.one("SELECT COALESCE(SUM(amount),0) tt FROM expenses WHERE expense_date BETWEEN ? AND ?", (a, b))
            marge = ts["tt"] - tp["tt"] - te["tt"]
            summary = [["Total Ventes", round(ts["tt"], 2)], ["Total Encaissé", round(ts["pp"], 2)],
                       ["Crédits Clients", round(ts["tt"] - ts["pp"], 2)],
                       ["Total Achats", round(tp["tt"], 2)], ["Total Dépenses", round(te["tt"], 2)],
                       ["Résultat Net", round(marge, 2)]]
            self._write_xlsx(path, [
                ("Mouvements", mheads, mvt, CURRENT_LANG == "ar"),
                ("Résumé", ["Indicateur", "Montant DA"], summary, False)])
            beep_ok()
            messagebox.showinfo(t("app_title"),
                                "تم تصدير التقرير إلى إكسل بنجاح !" if CURRENT_LANG == "ar"
                                else "Rapport exporté vers Excel avec succès !")
        except Exception as e:
            beep_err()
            messagebox.showerror(t("app_title"), f"Erreur Export Excel : {e}")

    # ==========================================
    # 5. المشتريات واستيراد Excel
    # ==========================================
    def show_purchases(self):
        self.clear()
        self.heading(t("purchases"), t("subtitle"))
        bar = tk.Frame(self.body, bg=BG)
        bar.pack(fill="x", pady=(0, 8))
        side = "left" if CURRENT_LANG=="fr" else "right"
        tk.Button(bar, text=t("new_purchase"), command=self.new_purchase, bg=GOLD, fg=BROWN,
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=12, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("import_excel"), command=self.import_excel_modal, bg="#217A4B", fg="white",
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=11, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("download_template"), command=self.save_excel_template, bg="#E8DFD1", fg=BROWN,
                  relief="flat", font=("Segoe UI", 9), padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("edit_purchase"), command=self.edit_purchase_dialog, bg="#E8DFD1", fg=BROWN,
                  relief="flat", font=("Segoe UI", 9), padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("delete_purchase"), command=self.delete_purchase_action, bg="#F7D8D5", fg=RED,
                  relief="flat", font=("Segoe UI", 9), padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("print_purchases_list"), command=self.print_purchases_report, bg="#E8DFD1", fg=BROWN,
                  relief="flat", font=("Segoe UI", 9), padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("export_excel"), command=self.export_purchases_excel, bg="#217A4B", fg="white",
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)

        search_frame = tk.Frame(bar, bg=BG)
        search_frame.pack(side="right" if CURRENT_LANG=="fr" else "left", padx=5)
        tk.Label(search_frame, text=f"🔍 {t('search')}:", bg=BG, font=("Segoe UI", 9, "bold")).pack(side="left", padx=3)
        self.purch_search = tk.Entry(search_frame, font=("Segoe UI", 9), width=16)
        self.purch_search.pack(side="left", padx=3, ipady=3)
        self.purch_search.bind("<KeyRelease>", lambda e: self.filter_purchases())

        cols = ("id", "date", "product", "supplier", "qty", "total", "paid", "due")
        heads = ("N°", t("date"), t("designation"), t("supplier"), t("quantity"), t("total"), t("paid"), t("due"))
        self.purchase_tree = self.create_tree(self.body, cols, heads, [55, 105, 270, 160, 75, 110, 110, 110], 9)
        self.purchase_tree.bind("<Double-1>", lambda e: self.edit_purchase_dialog())

        summary_bar = tk.Frame(self.body, bg="#EDE4D6", height=32, highlightbackground=BORDER, highlightthickness=1)
        summary_bar.pack(fill="x", pady=(6, 0))
        self.purch_summary_label = tk.Label(summary_bar, text="", bg="#EDE4D6", fg=BROWN, font=("Segoe UI", 9, "bold"))
        self.purch_summary_label.pack(side="left" if CURRENT_LANG=="fr" else "right", padx=15, pady=4)
        self.refresh_purchases()

    def filter_purchases(self):
        q = self.purch_search.get() if hasattr(self, "purch_search") else ""
        self.refresh_purchases(q)

    def refresh_purchases(self, query=""):
        if not hasattr(self, "purchase_tree") or not self.purchase_tree.winfo_exists(): return
        self.purchase_tree.delete(*self.purchase_tree.get_children())
        sql = """SELECT p.id, p.purchase_date, p.supplier, p.total, p.paid, (p.total - p.paid) AS due, p.notes,
                        COUNT(pi.id) AS items_count,
                        COALESCE(SUM(pi.quantity), 0) AS total_qty,
                        MIN(pi.product_name) AS sample_product
                 FROM purchases p
                 LEFT JOIN purchase_items pi ON pi.purchase_id = p.id
                 GROUP BY p.id
                 ORDER BY p.id ASC"""
        rows = self.db.all(sql)
        if query:
            q = query.strip().lower()
            rows = [r for r in rows if q in str(r["id"]).lower() or q in str(r["supplier"] or '').lower() or q in str(r["purchase_date"]).lower() or q in str(r["sample_product"] or '').lower() or q in str(r["notes"] or '').lower()]
        for r in rows:
            cnt = r["items_count"]
            if cnt == 1:
                p_name = display_product_name(r["sample_product"])
            elif cnt > 1:
                p_name = (f"📦 ({cnt} منتجات): " + display_product_name(r["sample_product"]) + " ...") if CURRENT_LANG=="ar" else f"📦 ({cnt} produits) : {r['sample_product']} ..."
            else:
                p_name = r["notes"] or "-"
            q_disp = f"{r['total_qty']:g}" if r["total_qty"] else "0"
            self.purchase_tree.insert("", "end", values=(
                r["id"], r["purchase_date"], p_name, r["supplier"],
                q_disp, money(r["total"]), money(r["paid"]), money(r["due"])
            ))
        if hasattr(self, "purch_summary_label") and self.purch_summary_label.winfo_exists():
            tot_purch = sum(r["total"] for r in rows)
            tot_paid = sum(r["paid"] for r in rows)
            tot_due = sum(r["due"] for r in rows)
            tot_q = sum(r["total_qty"] for r in rows)
            lbl = f"📦 إجمالي فواتير الشراء: {len(rows)} | مجموع الكميات: {tot_q:g} | المجموع: {money(tot_purch)} | المدفوع: {money(tot_paid)} | الديون: {money(tot_due)}" if CURRENT_LANG=="ar" else f"📦 Factures d'Achat : {len(rows)} | Quantités : {tot_q:g} | Total : {money(tot_purch)} | Payé : {money(tot_paid)} | Dettes : {money(tot_due)}"
            self.purch_summary_label.config(text=lbl)

    def selected_purchase_id(self):
        sel = self.purchase_tree.selection()
        if not sel:
            messagebox.showwarning(t("app_title"), "Sélectionnez un achat" if CURRENT_LANG=="fr" else "يرجى تحديد فاتورة شراء")
            return None
        return int(self.purchase_tree.item(sel[0], "values")[0])

    def edit_purchase_dialog(self):
        pid = self.selected_purchase_id()
        if not pid: return
        p = self.db.one("SELECT * FROM purchases WHERE id=?", (pid,))
        items = self.db.all("SELECT * FROM purchase_items WHERE purchase_id=?", (pid,))
        win = tk.Toplevel(self)
        win.title(f"{t('edit_purchase')} N° {pid}")
        center_window(win, 620, 390)
        win.configure(bg=BG)
        bottom_bar = tk.Frame(win, bg="#EDE4D6", height=45)
        bottom_bar.pack(side="bottom", fill="x")
        tk.Label(win, text=f"Achat N° {pid} - Total: {money(p['total'])}", bg=BG, fg=BROWN, font=("Segoe UI", 11, "bold")).pack(pady=6)
        f = tk.Frame(win, bg=BG); f.pack(padx=15, pady=2)
        tk.Label(f, text=t("supplier"), bg=BG, font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=8, pady=3, sticky="e")
        supp_entry = tk.Entry(f, width=22, font=("Segoe UI", 9)); supp_entry.grid(row=0, column=1, padx=8, pady=3)
        supp_entry.insert(0, p["supplier"])
        tk.Label(f, text=t("date"), bg=BG, font=("Segoe UI", 9, "bold")).grid(row=0, column=2, padx=8, pady=3, sticky="e")
        dt_entry = tk.Entry(f, width=15, font=("Segoe UI", 9)); dt_entry.grid(row=0, column=3, padx=8, pady=3)
        dt_entry.insert(0, p["purchase_date"])
        tk.Label(f, text=t("paid"), bg=BG, font=("Segoe UI", 9, "bold")).grid(row=1, column=0, padx=8, pady=3, sticky="e")
        paid_entry = tk.Entry(f, width=22, font=("Segoe UI", 9)); paid_entry.grid(row=1, column=1, padx=8, pady=3)
        paid_entry.insert(0, str(p["paid"]))

        if items:
            items_box = tk.LabelFrame(win, text="📦 السلع المشتراة" if CURRENT_LANG=="ar" else "📦 Articles achetés", bg=BG, fg=BROWN, font=("Segoe UI", 9, "bold"), padx=6, pady=4)
            items_box.pack(fill="both", expand=True, padx=15, pady=4)
            it_tree = ttk.Treeview(items_box, columns=("name", "qty", "cost", "subtot"), show="headings", height=4)
            it_tree.heading("name", text=t("designation"))
            it_tree.heading("qty", text=t("quantity"))
            it_tree.heading("cost", text=t("unit_cost"))
            it_tree.heading("subtot", text=t("total"))
            it_tree.column("name", width=250)
            it_tree.column("qty", width=70, anchor="center")
            it_tree.column("cost", width=90, anchor="center")
            it_tree.column("subtot", width=95, anchor="center")
            it_tree.pack(fill="both", expand=True)
            for it in items:
                it_tree.insert("", "end", values=(display_product_name(it["product_name"]), f"{it['quantity']:g}", money(it["unit_cost"]), money(it["subtotal"])))

        def save():
            pd = num(paid_entry.get())
            if pd < 0 or pd > p["total"]:
                beep_err(); messagebox.showerror(t("app_title"), "Montant payé invalide", parent=win); return
            if not messagebox.askyesno(t("app_title"), f"هل أنت متأكد من حفظ تعديلات الشراء رقم {pid}؟" if CURRENT_LANG=="ar" else "Confirmer l'enregistrement ?"):
                return
            self.db.run("UPDATE purchases SET supplier=?, purchase_date=?, paid=? WHERE id=?",
                        (supp_entry.get().strip(), dt_entry.get().strip(), pd, pid))
            beep_ok()
            win.destroy()
            self.refresh_purchases()
            messagebox.showinfo(t("app_title"), t("saved_success"))

        tk.Button(bottom_bar, text="💾 حفظ التعديلات" if CURRENT_LANG=="ar" else "💾 Enregistrer",
                  command=save, bg=GOLD, fg=BROWN, relief="flat", font=("Segoe UI", 9, "bold"), padx=16, pady=4).pack(pady=7)

    def delete_purchase_action(self):
        pid = self.selected_purchase_id()
        if not pid: return
        if not messagebox.askyesno(t("app_title"), t("purchase_delete_warning")): return
        items = self.db.all("SELECT product_id, quantity FROM purchase_items WHERE purchase_id=?", (pid,))
        for it in items:
            if it["product_id"]:
                self.db.run("UPDATE products SET stock = stock - ? WHERE id = ?", (it["quantity"], it["product_id"]))
        self.db.run("DELETE FROM purchases WHERE id=?", (pid,))
        self.db.resequence_table("purchases")
        beep_ok()
        self.refresh_purchases()
        messagebox.showinfo(t("app_title"), t("saved_success"))

    def print_purchases_report(self):
        sql = """SELECT p.id, p.purchase_date, p.supplier, p.total, p.paid, (p.total - p.paid) AS due,
                        COUNT(pi.id) AS items_count,
                        COALESCE(SUM(pi.quantity), 0) AS total_qty,
                        MIN(pi.product_name) AS sample_product
                 FROM purchases p
                 LEFT JOIN purchase_items pi ON pi.purchase_id = p.id
                 GROUP BY p.id
                 ORDER BY p.id ASC"""
        rows = self.db.all(sql)
        data = []
        for r in rows:
            cnt = r["items_count"]
            p_name = display_product_name(r["sample_product"]) if cnt == 1 else (f"({cnt} articles) {r['sample_product']}..." if cnt > 1 else "-")
            data.append([r["id"], r["purchase_date"], p_name, r["supplier"], f"{r['total_qty']:g}", money(r["total"]), money(r["paid"]), money(r["due"])])
        headers = ["N°", t("date"), t("designation"), t("supplier"), t("quantity"), t("total"), t("paid"), t("due")]
        self.print_table_report(t("print_purchases_list"), headers, data, "journal_achats")

    def new_purchase(self):
        products = self.db.all("SELECT * FROM products ORDER BY name")
        if not products:
            messagebox.showwarning(t("app_title"), "يرجى إضافة منتجات أولاً" if CURRENT_LANG=="ar" else "Ajoutez d'abord des produits")
            return
        win = tk.Toplevel(self)
        win.title(t("new_purchase"))
        center_window(win, 740, 500)
        win.configure(bg=BG)
        bottom_bar = tk.Frame(win, bg="#EDE4D6", height=48)
        bottom_bar.pack(side="bottom", fill="x")
        tk.Label(win, text=t("new_purchase"), bg=BG, fg=BROWN, font=("Segoe UI", 13, "bold")).pack(pady=6)
        f = tk.Frame(win, bg=BG); f.pack(fill="x", padx=12, pady=2)
        tk.Label(f, text=t("supplier"), bg=BG, font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=4, pady=3, sticky="e")
        supp_entry = tk.Entry(f, width=24, font=("Segoe UI", 9)); supp_entry.grid(row=0, column=1, padx=4)
        tk.Label(f, text=t("date"), bg=BG, font=("Segoe UI", 9, "bold")).grid(row=0, column=2, padx=4, sticky="e")
        dt_entry = tk.Entry(f, width=12, justify="center", font=("Segoe UI", 9))
        dt_entry.insert(0, today()); dt_entry.grid(row=0, column=3, padx=4)
        pmap = {f"{p['id']} - {p['name']}": p for p in products}
        tk.Label(f, text=t("products"), bg=BG, font=("Segoe UI", 9, "bold")).grid(row=1, column=0, padx=4, pady=3, sticky="e")
        prod_combo = ttk.Combobox(f, values=list(pmap), state="readonly", width=23)
        prod_combo.current(0); prod_combo.grid(row=1, column=1, padx=4)
        tk.Label(f, text=t("quantity"), bg=BG, font=("Segoe UI", 9, "bold")).grid(row=1, column=2, padx=4, sticky="e")
        q_entry = tk.Entry(f, width=12, justify="center", font=("Segoe UI", 9))
        q_entry.insert(0, "1"); q_entry.grid(row=1, column=3, padx=4)
        tk.Label(f, text=t("unit_cost"), bg=BG, font=("Segoe UI", 9, "bold")).grid(row=2, column=0, padx=4, pady=3, sticky="e")
        cost_entry = tk.Entry(f, width=12, justify="center", font=("Segoe UI", 9)); cost_entry.grid(row=2, column=1, padx=4, sticky="w")

        def set_cost(_=None):
            p = pmap.get(prod_combo.get())
            cost_entry.delete(0, "end"); cost_entry.insert(0, str(p["purchase_price"] if p else 0))
        prod_combo.bind("<<ComboboxSelected>>", set_cost); set_cost()

        items = []
        tree = self.create_tree(win, ("name", "qty", "cost", "total"),
                                (t("designation"), t("quantity"), t("unit_cost"), t("total")), [270, 80, 110, 120], 5)
        total_v = tk.StringVar(value=money(0))

        def add_item():
            p = pmap.get(prod_combo.get()); q = num(q_entry.get()); c = num(cost_entry.get())
            if not p or q <= 0 or c < 0:
                beep_err(); messagebox.showerror(t("app_title"), t("field_required"), parent=win); return
            items.append({"p": p, "qty": q, "cost": c})
            tree.insert("", "end", values=(p["name"], q, money(c), money(q * c)))
            total_v.set(money(sum(x["qty"] * x["cost"] for x in items)))
            paid_entry.delete(0, "end"); paid_entry.insert(0, str(sum(x["qty"] * x["cost"] for x in items)))

        tk.Button(f, text="إضافة" if CURRENT_LANG=="ar" else "Ajouter", command=add_item,
                  bg="#E8DFD1", font=("Segoe UI", 9, "bold"), relief="flat", padx=12).grid(row=2, column=2, pady=4)
        b = tk.Frame(win, bg=BG); b.pack(fill="x", padx=15, pady=4)
        tk.Label(b, text=f"{t('total')}: ", bg=BG, font=("Segoe UI", 10, "bold")).pack(side="left")
        tk.Label(b, textvariable=total_v, bg=BG, fg=BROWN, font=("Segoe UI", 11, "bold")).pack(side="left")
        tk.Label(b, text=f"{t('paid')}:", bg=BG, font=("Segoe UI", 9, "bold")).pack(side="left", padx=(15, 4))
        paid_entry = tk.Entry(b, width=12, justify="center", font=("Segoe UI", 9)); paid_entry.insert(0, "0"); paid_entry.pack(side="left")

        def save():
            if not items:
                beep_err(); messagebox.showwarning(t("app_title"), t("field_required"), parent=win); return
            tot = sum(x["qty"] * x["cost"] for x in items); pd = num(paid_entry.get())
            cur = self.db.conn.cursor()
            cur.execute("INSERT INTO purchases(purchase_date, supplier, total, paid) VALUES(?, ?, ?, ?)",
                        (dt_entry.get() or today(), supp_entry.get(), tot, pd))
            pid = cur.lastrowid
            for it in items:
                p = it["p"]
                cur.execute("""INSERT INTO purchase_items(purchase_id, product_id, product_name, quantity, unit_cost, subtotal)
                               VALUES(?, ?, ?, ?, ?, ?)""", (pid, p["id"], p["name"], it["qty"], it["cost"], it["qty"] * it["cost"]))
                cur.execute("UPDATE products SET stock=stock+?, purchase_price=? WHERE id=?", (it["qty"], it["cost"], p["id"]))
            self.db.conn.commit()
            beep_ok()
            win.destroy()
            self.refresh_purchases()
            messagebox.showinfo(t("app_title"), t("saved_success"))

        tk.Button(bottom_bar, text=f"💾 {t('new_purchase')}", command=save, bg=GOLD, fg=BROWN,
                  font=("Segoe UI", 10, "bold"), relief="flat", padx=18, pady=5, cursor="hand2").pack(pady=7)

    # --- استيراد Excel ---
    def save_excel_template(self):
        dest = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile="modele_import_achats.xlsx"
        )
        if not dest: return
        try:
            if not OPENPYXL_AVAILABLE:
                messagebox.showerror(t("app_title"), "openpyxl non installé.")
                return
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill, Alignment
            wb = Workbook()
            ws = wb.active
            ws.title = "Produits_Achats"
            headers = ["Désignation", "Quantité", "Prix Achat", "Prix Vente", "Catégorie", "Code-barres", "Charg", "Unité", "Date Expiration", "Fournisseur"]
            ws.append(headers)
            sample_rows = [
                ["Biscuit Bimo Chocolat 85g", 120, 35.0, 45.0, "Biscuits", "6130001001", 24, "Carton", "2027-12-31", "Bimo Algérie"],
                ["Gaufrette Maxi Vanille 50g", 240, 25.0, 35.0, "Gaufrettes", "6130001002", 48, "Paquet", "2026-10-15", "Saveurs d'Or"],
                ["Bonbon Fraise Acidulé", 50, 180.0, 230.0, "Confiserie", "6130001003", 12, "Boîte", "2028-06-30", "Sarl Trafik"]
            ]
            for row in sample_rows:
                ws.append(row)
            header_fill = PatternFill(start_color="3B1B12", end_color="3B1B12", fill_type="solid")
            header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
            center_align = Alignment(horizontal="center", vertical="center")
            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 14)
            wb.save(dest)
            beep_ok()
            messagebox.showinfo(t("app_title"), "تم حفظ نموذج الإكسل الجاهز بنجاح !" if CURRENT_LANG=="ar" else "Le modèle Excel a été enregistré avec succès !")
        except Exception as e:
            messagebox.showerror(t("app_title"), f"Erreur: {e}")

    def import_excel_modal(self):
        filepath = filedialog.askopenfilename(
            title="Sélectionner le fichier Excel / اختر ملف إكسل",
            filetypes=[("Excel & CSV", "*.xlsx *.xls *.csv"), ("Tous les fichiers", "*.*")]
        )
        if not filepath: return
        items = []
        try:
            if filepath.lower().endswith(".csv"):
                with open(filepath, "r", encoding="utf-8-sig") as f:
                    reader = csv.reader(f, delimiter=";")
                    rows = list(reader)
            else:
                if not OPENPYXL_AVAILABLE:
                    messagebox.showerror(t("app_title"), "openpyxl non installé."); return
                wb = openpyxl.load_workbook(filepath, data_only=True)
                ws = wb.active
                rows = list(ws.iter_rows(values_only=True))
                def is_product_row(h_row):
                    if not h_row: return False
                    h_lower = [str(c or '').strip().lower() for c in h_row]
                    return any(any(k in h for k in ['nom', 'désignation', 'designation', 'produit', 'المنتج', 'اسم', 'السلعة']) for h in h_lower)
                if not rows or len(rows) < 2 or not is_product_row(rows[0]):
                    for sname in wb.sheetnames:
                        c_rows = list(wb[sname].iter_rows(values_only=True))
                        if c_rows and len(c_rows) >= 2 and is_product_row(c_rows[0]):
                            ws = wb[sname]
                            rows = c_rows
                            break
            if not rows or len(rows) < 2:
                messagebox.showerror(t("app_title"), "Le fichier sélectionné est vide."); return
            header_row = [str(c or '').strip().lower() for c in rows[0]]
            col_map = {}
            for idx, h in enumerate(header_row):
                if any(k in h for k in ['fournisseur', 'مورد']):
                    col_map['supplier'] = idx
                elif any(k in h for k in ['prix achat', 'prix_achat', 'achat', 'cout', 'coût', 'شراء', 'تكلفة']):
                    col_map['buy_price'] = idx
                elif any(k in h for k in ['prix vente', 'prix_vente', 'vente', 'بيع']):
                    col_map['sell_price'] = idx
                elif any(k in h for k in ['quantité', 'quantite', 'qte', 'qté', 'الكمية', 'عدد', 'qte_achat']):
                    col_map['qty'] = idx
                elif any(k in h for k in ['catégorie', 'categorie', 'صنف', 'فئة', 'نوع']):
                    col_map['category'] = idx
                elif any(k in h for k in ['code-barres', 'code barres', 'barcode', 'باركود', 'كود']):
                    col_map['barcode'] = idx
                elif any(k in h for k in ['charg', 'حمولة', 'كرتون', 'colis']):
                    col_map['charg'] = idx
                elif any(k in h for k in ['unité', 'unite', 'الوحدة']):
                    col_map['unit'] = idx
                elif any(k in h for k in ['expiration', 'صلاحية', 'peremption']):
                    col_map['expiry'] = idx
                elif any(k in h for k in ['nom', 'désignation', 'designation', 'produit', 'المنتج', 'سلعة', 'بضاعة', 'اسم']):
                    col_map['name'] = idx
            if 'name' not in col_map:
                messagebox.showerror(t("app_title"), "Colonne 'Désignation / Nom' introuvable dans le fichier Excel."); return
            detected_supplier = ""
            for r in rows[1:]:
                if not r or all(c is None or str(c).strip() == '' for c in r):
                    continue
                name_val = str(r[col_map['name']] or '').strip()
                if not name_val:
                    continue
                qty_val = 1.0
                if 'qty' in col_map and r[col_map['qty']] is not None:
                    try: qty_val = float(str(r[col_map['qty']]).replace(',', '.').strip())
                    except: qty_val = 1.0
                buy_val = 0.0
                if 'buy_price' in col_map and r[col_map['buy_price']] is not None:
                    try: buy_val = float(str(r[col_map['buy_price']]).replace(',', '.').strip())
                    except: buy_val = 0.0
                sell_val = round(buy_val * 1.15, 2)
                if 'sell_price' in col_map and r[col_map['sell_price']] is not None:
                    try: sell_val = float(str(r[col_map['sell_price']]).replace(',', '.').strip())
                    except: pass
                cat_val = str(r[col_map['category']] or '').strip() if 'category' in col_map and r[col_map['category']] else ''
                code_val = str(r[col_map['barcode']] or '').strip() if 'barcode' in col_map and r[col_map['barcode']] else ''
                charg_val = float(r[col_map['charg']]) if 'charg' in col_map and r[col_map['charg']] is not None else 1.0
                unit_val = str(r[col_map['unit']] or 'Boîte').strip() if 'unit' in col_map and r[col_map['unit']] else 'Boîte'
                expiry_val = str(r[col_map['expiry']] or '').strip() if 'expiry' in col_map and r[col_map['expiry']] else ''
                supp_val = str(r[col_map['supplier']] or '').strip() if 'supplier' in col_map and r[col_map['supplier']] else ''
                if supp_val and not detected_supplier:
                    detected_supplier = supp_val
                items.append({
                    'name': name_val, 'qty': qty_val, 'buy_price': buy_val,
                    'sell_price': sell_val, 'category': cat_val, 'barcode': code_val,
                    'charg': charg_val, 'unit': unit_val, 'expiry': expiry_val, 'supplier': supp_val
                })
        except Exception as e:
            messagebox.showerror(t("app_title"), f"Erreur lors de la lecture du fichier : {e}")
            return
        if not items:
            messagebox.showwarning(t("app_title"), "Aucun produit valide trouvé dans ce fichier."); return

        win = tk.Toplevel(self)
        win.title("Prévisualisation Import Excel / استيراد المشتريات والمخزون")
        center_window(win, 880, 560)
        win.configure(bg=BG)
        bottom_bar = tk.Frame(win, bg="#EDE4D6", height=50)
        bottom_bar.pack(side="bottom", fill="x")
        header = tk.Frame(win, bg=HEADER_BG, height=45)
        header.pack(fill="x")
        tk.Label(header, text=f"📥 Importation Excel ({len(items)} produits détectés)", bg=HEADER_BG, fg=GOLD, font=("Segoe UI", 12, "bold")).pack(pady=10)
        params_frame = tk.Frame(win, bg=BG)
        params_frame.pack(fill="x", padx=15, pady=8)
        tk.Label(params_frame, text=f"{t('supplier')}:", bg=BG, font=("Segoe UI", 9, "bold")).pack(side="left", padx=4)
        supp_entry = tk.Entry(params_frame, font=("Segoe UI", 9), width=22)
        supp_entry.insert(0, detected_supplier or "Fournisseur Import Excel")
        supp_entry.pack(side="left", padx=4)
        tk.Label(params_frame, text=f"{t('date')}:", bg=BG, font=("Segoe UI", 9, "bold")).pack(side="left", padx=(12, 4))
        date_entry = tk.Entry(params_frame, font=("Segoe UI", 9), width=12, justify="center")
        date_entry.insert(0, today())
        date_entry.pack(side="left", padx=4)
        tot_purchase = sum(it["qty"] * it["buy_price"] for it in items)
        tot_qty = sum(it["qty"] for it in items)
        tk.Label(params_frame, text=f"{t('paid')}:", bg=BG, font=("Segoe UI", 9, "bold")).pack(side="left", padx=(12, 4))
        paid_entry = tk.Entry(params_frame, font=("Segoe UI", 9), width=14, justify="center")
        paid_entry.insert(0, str(tot_purchase))
        paid_entry.pack(side="left", padx=4)

        prev_cols = ("num", "name", "category", "barcode", "qty", "buy", "sell", "subtotal")
        prev_heads = ("N°", "Désignation", "Catégorie", "Code-barres", "Quantité", "Prix Achat", "Prix Vente", "Total Achat")
        prev_tree = self.create_tree(win, prev_cols, prev_heads, [40, 240, 90, 110, 70, 95, 95, 110], 8)
        for idx, it in enumerate(items, start=1):
            prev_tree.insert("", "end", values=(
                idx, it["name"], it["category"] or "-", it["barcode"] or "-",
                f"{it['qty']:g}", money(it["buy_price"]), money(it["sell_price"]), money(it["qty"] * it["buy_price"])
            ))
        sum_label = tk.Label(win, text=f"Total Articles: {len(items)}  |  Total Quantités: {tot_qty:g}  |  Montant Total Achat: {money(tot_purchase)}",
                             bg=BG, fg=BROWN, font=("Segoe UI", 10, "bold"))
        sum_label.pack(pady=4)

        def confirm_import():
            supp_name = supp_entry.get().strip() or "Fournisseur Import Excel"
            p_date = date_entry.get().strip() or today()
            p_paid = num(paid_entry.get())
            try:
                cur = self.db.conn.cursor()
                cur.execute("INSERT INTO purchases(purchase_date, supplier, total, paid, notes) VALUES(?, ?, ?, ?, ?)",
                            (p_date, supp_name, tot_purchase, p_paid, f"Importé depuis {Path(filepath).name}"))
                pid = cur.lastrowid
                for it in items:
                    existing = None
                    if it["barcode"]:
                        existing = self.db.one("SELECT * FROM products WHERE barcode=?", (it["barcode"],))
                    if not existing:
                        existing = self.db.one("SELECT * FROM products WHERE name=?", (it["name"],))
                    if existing:
                        prod_id = existing["id"]
                        cur.execute("""UPDATE products SET stock = stock + ?, purchase_price = ?,
                                       sale_price = CASE WHEN ? > 0 THEN ? ELSE sale_price END
                                       WHERE id = ?""",
                                    (it["qty"], it["buy_price"], it["sell_price"], it["sell_price"], prod_id))
                    else:
                        cur.execute("""INSERT INTO products(name, category, barcode, charg, purchase_price, sale_price, stock, min_stock, unit, expiry, created_at)
                                       VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                                    (it["name"], it["category"], it["barcode"], it["charg"], it["buy_price"], it["sell_price"],
                                     it["qty"], 5, it["unit"], it["expiry"], today()))
                        prod_id = cur.lastrowid
                    cur.execute("""INSERT INTO purchase_items(purchase_id, product_id, product_name, quantity, unit_cost, subtotal)
                                   VALUES(?, ?, ?, ?, ?, ?)""",
                                (pid, prod_id, it["name"], it["qty"], it["buy_price"], it["qty"] * it["buy_price"]))
                self.db.conn.commit()
            except Exception as e:
                self.db.conn.rollback()
                beep_err()
                messagebox.showerror(t("app_title"), f"Erreur lors de l'enregistrement : {e}", parent=win)
                return
            self.db.resequence_table("purchases")
            self.db.resequence_table("products")
            beep_ok()
            win.destroy()
            self.refresh_purchases()
            self.refresh_products()
            success_txt = f"تم استيراد {len(items)} منتجاً بنجاح!\nإجمالي المشتريات: {money(tot_purchase)}" if CURRENT_LANG=="ar" else f"{len(items)} produits importés avec succès !\nMontant total: {money(tot_purchase)}"
            messagebox.showinfo(t("app_title"), success_txt)

        btn_txt = "✅ تأكيد الاستيراد والإدخال المباشر للمخزن" if CURRENT_LANG=="ar" else "✅ Confirmer l'Importation vers le Stock"
        tk.Button(bottom_bar, text=btn_txt, command=confirm_import, bg="#217A4B", fg="white",
                  font=("Segoe UI", 10, "bold"), relief="flat", padx=20, pady=5, cursor="hand2").pack(pady=7)

    # ==========================================
    # 6. المصاريف
    # ==========================================
    def show_expenses(self):
        self.clear()
        self.heading(t("expenses"), t("subtitle"))
        bar = tk.Frame(self.body, bg=BG)
        bar.pack(fill="x", pady=(0, 8))
        side = "left" if CURRENT_LANG=="fr" else "right"
        tk.Button(bar, text=t("add_expense"), command=self.add_expense, bg=GOLD, fg=BROWN,
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=12, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("edit_expense"), command=self.edit_expense_dialog, bg="#E8DFD1", fg=BROWN,
                  relief="flat", font=("Segoe UI", 9), padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("delete_expense"), command=self.delete_expense_action, bg="#F7D8D5", fg=RED,
                  relief="flat", font=("Segoe UI", 9), padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("print_expenses_list"), command=self.print_expenses_report, bg="#E8DFD1", fg=BROWN,
                  relief="flat", font=("Segoe UI", 9), padx=10, pady=6, cursor="hand2").pack(side=side, padx=3)

        cols = ("id", "date", "title", "amount", "notes")
        heads = ("N°", t("date"), t("expense_title"), t("amount"), t("notes"))
        self.expense_tree = self.create_tree(self.body, cols, heads, [50, 110, 240, 130, 300], 9)
        self.expense_tree.bind("<Double-1>", lambda e: self.edit_expense_dialog())
        self.refresh_expenses()

    def refresh_expenses(self):
        if not hasattr(self, "expense_tree") or not self.expense_tree.winfo_exists(): return
        self.expense_tree.delete(*self.expense_tree.get_children())
        for r in self.db.all("SELECT * FROM expenses ORDER BY id ASC"):
            self.expense_tree.insert("", "end", values=(
                r["id"], r["expense_date"], r["title"], money(r["amount"]), r["notes"]
            ))

    def selected_expense_id(self):
        sel = self.expense_tree.selection()
        if not sel:
            messagebox.showwarning(t("app_title"), "Sélectionnez une dépense" if CURRENT_LANG=="fr" else "يرجى تحديد مصروف")
            return None
        return int(self.expense_tree.item(sel[0], "values")[0])

    def add_expense(self):
        self.expense_dialog()

    def edit_expense_dialog(self):
        eid = self.selected_expense_id()
        if eid:
            row = self.db.one("SELECT * FROM expenses WHERE id=?", (eid,))
            self.expense_dialog(row)

    def expense_dialog(self, row=None):
        win = tk.Toplevel(self)
        win.title(t("edit_expense") if row else t("add_expense"))
        center_window(win, 440, 260)
        win.configure(bg=BG)
        bottom_bar = tk.Frame(win, bg="#EDE4D6", height=44)
        bottom_bar.pack(side="bottom", fill="x")
        labels = [t("date"), t("expense_title"), t("amount"), t("notes")]
        keys = ["expense_date", "title", "amount", "notes"]
        entries = {}
        form = tk.Frame(win, bg=BG)
        form.pack(fill="both", expand=True, padx=15, pady=8)
        for i, (l, k) in enumerate(zip(labels, keys)):
            tk.Label(form, text=l, bg=BG, fg=BROWN, font=("Segoe UI", 9, "bold")).grid(
                row=i, column=0, padx=8, pady=4, sticky="e" if CURRENT_LANG=="fr" else "w")
            e = tk.Entry(form, width=28, font=("Segoe UI", 9))
            e.grid(row=i, column=1, padx=8, pady=4, ipady=2)
            if row: e.insert(0, str(row[k] or ""))
            elif k == "expense_date": e.insert(0, today())
            entries[k] = e

        def save():
            title = entries["title"].get().strip()
            amt = num(entries["amount"].get())
            if not title or amt <= 0:
                beep_err(); messagebox.showerror(t("app_title"), t("field_required"), parent=win); return
            if row:
                self.db.run("UPDATE expenses SET expense_date=?, title=?, amount=?, notes=? WHERE id=?",
                            (entries["expense_date"].get().strip(), title, amt, entries["notes"].get().strip(), row["id"]))
            else:
                self.db.run("INSERT INTO expenses(expense_date, title, amount, notes) VALUES(?, ?, ?, ?)",
                            (entries["expense_date"].get().strip(), title, amt, entries["notes"].get().strip()))
            beep_ok()
            win.destroy()
            self.refresh_expenses()

        tk.Button(bottom_bar, text="💾 حفظ البيانات" if CURRENT_LANG=="ar" else "💾 Enregistrer", command=save,
                  bg=GOLD, fg=BROWN, relief="flat", font=("Segoe UI", 9, "bold"), padx=18, pady=4).pack(pady=7)

    def delete_expense_action(self):
        eid = self.selected_expense_id()
        if not eid: return
        if messagebox.askyesno(t("app_title"), t("confirm_delete")):
            self.db.run("DELETE FROM expenses WHERE id=?", (eid,))
            self.db.resequence_table("expenses")
            beep_ok()
            self.refresh_expenses()
            messagebox.showinfo(t("app_title"), t("saved_success"))

    def print_expenses_report(self):
        rows = self.db.all("SELECT id, expense_date, title, amount, notes FROM expenses ORDER BY id ASC")
        data = [[r["id"], r["expense_date"], r["title"], money(r["amount"]), r["notes"] or "-"] for r in rows]
        headers = ["N°", t("date"), t("expense_title"), t("amount"), t("notes")]
        self.print_table_report(t("print_expenses_list"), headers, data, "journal_depenses")

    # ==========================================
    # 7. التقارير
    # ==========================================
    def show_reports(self):
        self.clear()
        self.heading(t("reports"), t("subtitle"))
        f = tk.Frame(self.body, bg=BG); f.pack(fill="x", pady=6)
        side = "left" if CURRENT_LANG=="fr" else "right"
        tk.Label(f, text=f"{t('from_date')}:", bg=BG, font=("Segoe UI", 9, "bold")).pack(side=side, padx=4)
        start_entry = tk.Entry(f, width=12, justify="center", font=("Segoe UI", 9))
        start_entry.insert(0, date.today().replace(day=1).isoformat()); start_entry.pack(side=side, padx=4, ipady=2)
        tk.Label(f, text=f"{t('to_date')}:", bg=BG, font=("Segoe UI", 9, "bold")).pack(side=side, padx=4)
        end_entry = tk.Entry(f, width=12, justify="center", font=("Segoe UI", 9))
        end_entry.insert(0, today()); end_entry.pack(side=side, padx=4, ipady=2)
        cards_out = tk.Frame(self.body, bg=BG); cards_out.pack(fill="both", expand=True, pady=10)

        def generate():
            a = start_entry.get().strip(); b = end_entry.get().strip()
            try:
                datetime.strptime(a, "%Y-%m-%d"); datetime.strptime(b, "%Y-%m-%d")
            except Exception:
                messagebox.showerror(t("app_title"), "Format de date YYYY-MM-DD requis"); return
            s = self.db.one("SELECT COALESCE(SUM(total),0) n, COALESCE(SUM(paid),0) p FROM sales WHERE sale_date BETWEEN ? AND ?", (a, b))
            pur = self.db.one("SELECT COALESCE(SUM(total),0) n FROM purchases WHERE purchase_date BETWEEN ? AND ?", (a, b))["n"]
            exp = self.db.one("SELECT COALESCE(SUM(amount),0) n FROM expenses WHERE expense_date BETWEEN ? AND ?", (a, b))["n"]
            cost = self.db.one("""SELECT COALESCE(SUM(si.quantity*si.unit_cost),0) n FROM sale_items si
                                  JOIN sales s ON s.id=si.sale_id WHERE s.sale_date BETWEEN ? AND ?""", (a, b))["n"]
            gross = s["n"] - cost
            net = gross - exp
            for w in cards_out.winfo_children(): w.destroy()
            stats = [
                (t("total_sales"), money(s["n"]), GOLD),
                (t("today_paid"), money(s["p"]), GREEN),
                (t("total_purchases"), money(pur), BLUE),
                (t("cogs"), money(cost), MUTED),
                (t("gross_margin"), money(gross), GREEN if gross >= 0 else RED),
                (t("total_expenses"), money(exp), RED),
                (t("net_profit"), money(net), GREEN if net >= 0 else RED)
            ]
            for i, (lab, val, col) in enumerate(stats):
                c = self.stat_card(cards_out, lab, val, col)
                c.grid(row=i//3, column=i%3, sticky="ew", padx=5, pady=5)
            for j in range(3): cards_out.columnconfigure(j, weight=1)

        tk.Button(f, text=f"🔍 {t('view_report')}", command=generate, bg=GOLD, fg=BROWN,
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=12, pady=4, cursor="hand2").pack(side=side, padx=8)
        tk.Button(f, text=f"📊 {t('export_csv')}", command=lambda: self.export_csv(start_entry.get(), end_entry.get()),
                  bg="#E8DFD1", fg=BROWN, relief="flat", font=("Segoe UI", 9), padx=10, pady=4, cursor="hand2").pack(side=side, padx=4)
        tk.Button(f, text=t("export_excel"), command=lambda: self.export_report_excel(start_entry.get(), end_entry.get()),
                  bg="#217A4B", fg="white", relief="flat", font=("Segoe UI", 9, "bold"), padx=10, pady=4, cursor="hand2").pack(side=side, padx=4)
        generate()

    def export_csv(self, a, b):
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")],
                                            initialfile=f"rapport_trafik_{a}_{b}.csv")
        if not path: return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Date", "Type", "Ref", "Total_DA", "Paye_DA", "Reste_DA"])
            for r in self.db.all("SELECT id, sale_date, total, paid FROM sales WHERE sale_date BETWEEN ? AND ?", (a, b)):
                w.writerow([r["sale_date"], "Vente", r["id"], r["total"], r["paid"], r["total"] - r["paid"]])
            for r in self.db.all("SELECT id, purchase_date, total, paid FROM purchases WHERE purchase_date BETWEEN ? AND ?", (a, b)):
                w.writerow([r["purchase_date"], "Achat", r["id"], r["total"], r["paid"], r["total"] - r["paid"]])
            for r in self.db.all("SELECT expense_date, title, amount FROM expenses WHERE expense_date BETWEEN ? AND ?", (a, b)):
                w.writerow([r["expense_date"], "Dépense", r["title"], r["amount"], r["amount"], 0])
        beep_ok()
        messagebox.showinfo(t("app_title"), t("saved_success"))

    # ==========================================
    # 8. الإعدادات
    # ==========================================
    def show_settings(self):
        if self.current_user.get("role") != "Admin":
            messagebox.showwarning(t("app_title"), "هذه الصفحة مخصصة للمسيرين فقط (Admin)" if CURRENT_LANG=="ar" else "Accès réservé aux Administrateurs")
            return
        self.clear()
        self.heading(t("settings"), "Gestion des accès, mots de passe et personnalisation de la facturation")
        notebook = ttk.Notebook(self.body)
        notebook.pack(fill="both", expand=True)
        tab_users = tk.Frame(notebook, bg=BG)
        tab_company = tk.Frame(notebook, bg=BG)
        notebook.add(tab_users, text=f"👥 {t('users_tab')}")
        notebook.add(tab_company, text=f"🏢 {t('company_tab')}")
        bar = tk.Frame(tab_users, bg=BG); bar.pack(fill="x", padx=8, pady=8)
        side = "left" if CURRENT_LANG=="fr" else "right"
        tk.Button(bar, text=t("add_user"), command=self.add_user_dialog, bg=GOLD, fg=BROWN,
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=10, pady=5, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("edit_user"), command=self.edit_user_dialog, bg="#E8DFD1", fg=BROWN,
                  relief="flat", font=("Segoe UI", 9), padx=10, pady=5, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("change_pwd"), command=self.change_user_password_dialog, bg="#E8DFD1", fg=BROWN,
                  relief="flat", font=("Segoe UI", 9), padx=10, pady=5, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("delete_user"), command=self.delete_user_action, bg="#F7D8D5", fg=RED,
                  relief="flat", font=("Segoe UI", 9), padx=10, pady=5, cursor="hand2").pack(side=side, padx=3)
        tk.Button(bar, text=t("print_users_list"), command=self.print_users_report, bg="#E8DFD1", fg=BROWN,
                  relief="flat", font=("Segoe UI", 9), padx=10, pady=5, cursor="hand2").pack(side=side, padx=3)
        u_cols = ("id", "username", "role", "created_at")
        u_heads = ("N°", t("username"), t("role"), "Date de création" if CURRENT_LANG=="fr" else "تاريخ الإنشاء")
        self.users_tree = self.create_tree(tab_users, u_cols, u_heads, [50, 180, 150, 160], 7)
        self.users_tree.bind("<Double-1>", lambda e: self.edit_user_dialog())
        self.refresh_users()

        c_frame = tk.Frame(tab_company, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
        c_frame.pack(fill="both", expand=True, padx=12, pady=10)
        fields = [
            ("company_name", t("company_name")),
            ("company_activity", t("company_activity")),
            ("company_phone", t("company_phone")),
            ("company_address", t("company_address")),
            ("company_rc", t("rc")),
            ("company_nif", t("nif")),
            ("invoice_footer", t("invoice_footer_msg")),
        ]
        self.comp_entries = {}
        for i, (k, lbl) in enumerate(fields):
            tk.Label(c_frame, text=lbl, bg=PANEL, fg=BROWN, font=("Segoe UI", 9, "bold")).grid(
                row=i, column=0, padx=15, pady=5, sticky="e" if CURRENT_LANG=="fr" else "w")
            e = tk.Entry(c_frame, font=("Segoe UI", 9), width=42)
            e.grid(row=i, column=1, padx=8, pady=5, ipady=3, sticky="w")
            e.insert(0, self.db.get_setting(k, ""))
            self.comp_entries[k] = e
        tk.Label(c_frame, text=t("invoice_format"), bg=PANEL, fg=BROWN, font=("Segoe UI", 9, "bold")).grid(
            row=len(fields), column=0, padx=15, pady=5, sticky="e" if CURRENT_LANG=="fr" else "w")
        self.fmt_combo = ttk.Combobox(c_frame, values=["A4", "Ticket"], state="readonly", width=14)
        self.fmt_combo.set(self.db.get_setting("invoice_format", "A4"))
        self.fmt_combo.grid(row=len(fields), column=1, padx=8, pady=5, sticky="w")

        def save_company_settings():
            for k, entry in self.comp_entries.items():
                self.db.set_setting(k, entry.get().strip())
            self.db.set_setting("invoice_format", self.fmt_combo.get())
            beep_ok()
            messagebox.showinfo(t("app_title"), t("saved_success"))

        tk.Button(c_frame, text=f"💾 {t('save_settings')}", command=save_company_settings,
                  bg=GOLD, fg=BROWN, relief="flat", font=("Segoe UI", 10, "bold"), padx=20, pady=7, cursor="hand2").grid(
            row=len(fields)+1, column=0, columnspan=2, pady=14)

    def refresh_users(self):
        if not hasattr(self, "users_tree") or not self.users_tree.winfo_exists(): return
        self.users_tree.delete(*self.users_tree.get_children())
        for u in self.db.all("SELECT id, username, role, created_at FROM users ORDER BY id ASC"):
            role_text = t("admin") if u["role"] == "Admin" else t("cashier")
            self.users_tree.insert("", "end", values=(u["id"], u["username"], role_text, u["created_at"]))

    def selected_user_id(self):
        sel = self.users_tree.selection()
        if not sel:
            messagebox.showwarning(t("app_title"), "Sélectionnez un utilisateur" if CURRENT_LANG=="fr" else "يرجى تحديد مستخدم من الجدول")
            return None
        return int(self.users_tree.item(sel[0], "values")[0])

    def add_user_dialog(self):
        win = tk.Toplevel(self)
        win.title(t("add_user"))
        center_window(win, 420, 260)
        win.configure(bg=BG)
        bottom_bar = tk.Frame(win, bg="#EDE4D6", height=42)
        bottom_bar.pack(side="bottom", fill="x")
        f = tk.Frame(win, bg=BG)
        f.pack(fill="both", expand=True, padx=15, pady=8)
        tk.Label(f, text=t("username"), bg=BG, font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=10, pady=5, sticky="e")
        u_entry = tk.Entry(f, width=22, font=("Segoe UI", 9)); u_entry.grid(row=0, column=1, padx=10, pady=5)
        tk.Label(f, text=t("password"), bg=BG, font=("Segoe UI", 9, "bold")).grid(row=1, column=0, padx=10, pady=5, sticky="e")
        p_entry = tk.Entry(f, width=22, show="•", font=("Segoe UI", 9)); p_entry.grid(row=1, column=1, padx=10, pady=5)
        tk.Label(f, text=t("confirm_pwd"), bg=BG, font=("Segoe UI", 9, "bold")).grid(row=2, column=0, padx=10, pady=5, sticky="e")
        cp_entry = tk.Entry(f, width=22, show="•", font=("Segoe UI", 9)); cp_entry.grid(row=2, column=1, padx=10, pady=5)
        tk.Label(f, text=t("role"), bg=BG, font=("Segoe UI", 9, "bold")).grid(row=3, column=0, padx=10, pady=5, sticky="e")
        role_combo = ttk.Combobox(f, values=["Admin", "Vendeur"], state="readonly", width=19)
        role_combo.current(0); role_combo.grid(row=3, column=1, padx=10, pady=5)

        def save():
            u = u_entry.get().strip(); p = p_entry.get().strip(); cp = cp_entry.get().strip()
            if not u or not p:
                beep_err(); messagebox.showerror(t("app_title"), t("field_required"), parent=win); return
            if p != cp:
                beep_err(); messagebox.showerror(t("app_title"), t("pwd_mismatch"), parent=win); return
            if self.db.one("SELECT id FROM users WHERE username=?", (u,)):
                beep_err(); messagebox.showerror(t("app_title"), "اسم المستخدم موجود مسبقاً" if CURRENT_LANG=="ar" else "Ce nom d'utilisateur existe déjà", parent=win); return
            self.db.run("INSERT INTO users(username, password, role, created_at) VALUES(?, ?, ?, ?)",
                        (u, p, role_combo.get(), today()))
            beep_ok()
            win.destroy()
            self.refresh_users()
            messagebox.showinfo(t("app_title"), t("saved_success"))

        tk.Button(bottom_bar, text="💾 حفظ" if CURRENT_LANG=="ar" else "💾 Enregistrer", command=save,
                  bg=GOLD, fg=BROWN, relief="flat", font=("Segoe UI", 9, "bold"), padx=16, pady=4).pack(pady=6)

    def edit_user_dialog(self):
        uid = self.selected_user_id()
        if not uid: return
        u_record = self.db.one("SELECT * FROM users WHERE id=?", (uid,))
        win = tk.Toplevel(self)
        win.title(f"{t('edit_user')} : {u_record['username']}")
        center_window(win, 420, 220)
        win.configure(bg=BG)
        bottom_bar = tk.Frame(win, bg="#EDE4D6", height=42)
        bottom_bar.pack(side="bottom", fill="x")
        f = tk.Frame(win, bg=BG)
        f.pack(fill="both", expand=True, padx=15, pady=8)
        tk.Label(f, text=t("username"), bg=BG, font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=10, pady=6, sticky="e")
        u_entry = tk.Entry(f, width=22, font=("Segoe UI", 9)); u_entry.grid(row=0, column=1, padx=10, pady=6)
        u_entry.insert(0, u_record["username"])
        tk.Label(f, text=t("role"), bg=BG, font=("Segoe UI", 9, "bold")).grid(row=1, column=0, padx=10, pady=6, sticky="e")
        role_combo = ttk.Combobox(f, values=["Admin", "Vendeur"], state="readonly", width=19)
        role_combo.set(u_record["role"])
        role_combo.grid(row=1, column=1, padx=10, pady=6)

        def save():
            u = u_entry.get().strip()
            if not u:
                beep_err(); messagebox.showerror(t("app_title"), t("field_required"), parent=win); return
            exists = self.db.one("SELECT id FROM users WHERE username=? AND id!=?", (u, uid))
            if exists:
                beep_err(); messagebox.showerror(t("app_title"), "اسم المستخدم موجود مسبقاً" if CURRENT_LANG=="ar" else "Ce nom d'utilisateur existe déjà", parent=win); return
            self.db.run("UPDATE users SET username=?, role=? WHERE id=?", (u, role_combo.get(), uid))
            beep_ok()
            win.destroy()
            self.refresh_users()
            messagebox.showinfo(t("app_title"), t("saved_success"))

        tk.Button(bottom_bar, text="💾 حفظ" if CURRENT_LANG=="ar" else "💾 Enregistrer", command=save,
                  bg=GOLD, fg=BROWN, relief="flat", font=("Segoe UI", 9, "bold"), padx=16, pady=4).pack(pady=6)

    def change_user_password_dialog(self):
        uid = self.selected_user_id()
        if not uid: return
        u_record = self.db.one("SELECT * FROM users WHERE id=?", (uid,))
        win = tk.Toplevel(self)
        win.title(f"{t('change_pwd')} : {u_record['username']}")
        center_window(win, 420, 220)
        win.configure(bg=BG)
        bottom_bar = tk.Frame(win, bg="#EDE4D6", height=42)
        bottom_bar.pack(side="bottom", fill="x")
        tk.Label(win, text=f"Utilisateur: {u_record['username']}", bg=BG, font=("Segoe UI", 10, "bold")).pack(pady=4)
        f = tk.Frame(win, bg=BG); f.pack(padx=15, pady=4)
        tk.Label(f, text=t("password"), bg=BG, font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=8, pady=4, sticky="e")
        p1 = tk.Entry(f, show="•", width=20, font=("Segoe UI", 9)); p1.grid(row=0, column=1, padx=8, pady=4)
        tk.Label(f, text=t("confirm_pwd"), bg=BG, font=("Segoe UI", 9, "bold")).grid(row=1, column=0, padx=8, pady=4, sticky="e")
        p2 = tk.Entry(f, show="•", width=20, font=("Segoe UI", 9)); p2.grid(row=1, column=1, padx=8, pady=4)

        def save():
            pwd = p1.get().strip()
            if not pwd:
                beep_err(); messagebox.showerror(t("app_title"), t("field_required"), parent=win); return
            if pwd != p2.get().strip():
                beep_err(); messagebox.showerror(t("app_title"), t("pwd_mismatch"), parent=win); return
            self.db.run("UPDATE users SET password=? WHERE id=?", (pwd, uid))
            beep_ok()
            win.destroy()
            messagebox.showinfo(t("app_title"), t("saved_success"))

        tk.Button(bottom_bar, text="تحديث كلمة المرور" if CURRENT_LANG=="ar" else "Mettre à jour",
                  command=save, bg=GOLD, fg=BROWN, relief="flat", font=("Segoe UI", 9, "bold"), padx=15, pady=4).pack(pady=6)

    def delete_user_action(self):
        uid = self.selected_user_id()
        if not uid: return
        if uid == self.current_user["id"]:
            beep_err(); messagebox.showerror(t("app_title"), t("cannot_delete_self")); return
        u_target = self.db.one("SELECT * FROM users WHERE id=?", (uid,))
        if u_target["role"] == "Admin":
            admin_count = self.db.one("SELECT COUNT(*) n FROM users WHERE role='Admin'")["n"]
            if admin_count <= 1:
                beep_err(); messagebox.showerror(t("app_title"), t("cannot_delete_last_admin")); return
        if messagebox.askyesno(t("app_title"), t("confirm_delete")):
            self.db.run("DELETE FROM users WHERE id=?", (uid,))
            self.db.resequence_table("users")
            beep_ok()
            self.refresh_users()

    def print_users_report(self):
        rows = self.db.all("SELECT id, username, role, created_at FROM users ORDER BY id ASC")
        data = [[r["id"], r["username"], t("admin") if r["role"]=="Admin" else t("cashier"), r["created_at"]] for r in rows]
        headers = ["N°", t("username"), t("role"), "Date de création" if CURRENT_LANG=="fr" else "تاريخ الإنشاء"]
        self.print_table_report(t("print_users_list"), headers, data, "liste_utilisateurs")

    # ==========================================
    # 9. النسخ الاحتياطي والإغلاق
    # ==========================================
    def backup(self):
        target = filedialog.asksaveasfilename(
            defaultextension=".db", initialdir=str(BACKUP_DIR),
            initialfile=f"trafik_saveurs_backup_{today()}.db",
            filetypes=[("SQLite Database", "*.db")]
        )
        if not target: return
        self.db.conn.commit()
        shutil.copy2(DB_PATH, target)
        beep_ok()
        messagebox.showinfo(t("backup"), t("saved_success"))

    def logout(self):
        msg = "هل تريد تسجيل الخروج حقاً؟" if CURRENT_LANG=="ar" else "Voulez-vous vraiment vous déconnecter ?"
        if messagebox.askyesno(t("logout"), msg):
            self.db.close()
            self.destroy()
            LoginApp().mainloop()

    def on_close(self):
        self.db.close()
        self.destroy()

# ============================================================
# نقطة البداية - تشغيل شاشة التحميل أولاً
# ============================================================
if __name__ == "__main__":
    splash = SplashScreen()
    splash.mainloop()
