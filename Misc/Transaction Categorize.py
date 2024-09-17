import csv
import tempfile
import shutil
import unicodedata

def clean_transaction_detail(transaction_detail):
    cleaned_detail = ''.join(char for char in unicodedata.normalize('NFKD', transaction_detail) if not unicodedata.combining(char))
    cleaned_detail = cleaned_detail.replace('Â', '').replace(' ', '').replace('POSDEB', '').replace('DBTCRD', '')
    return cleaned_detail

def extract_customer_name(transaction_detail, mappings):
    cleaned_detail = ''.join(transaction_detail.split()).upper()
    return next((customer_name for description, customer_name in mappings.items() if description in cleaned_detail), None)

def clean_csv_file(csv_file, mappings):
    with open(csv_file, 'r', newline='', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        fieldnames = reader.fieldnames

        with tempfile.NamedTemporaryFile(mode='w', delete=False, newline='', encoding='utf-8') as temp_file:
            writer = csv.DictWriter(temp_file, fieldnames=fieldnames)
            writer.writeheader()

            for row in reader:
                transaction_detail = row['Description']
                cleaned_detail = clean_transaction_detail(transaction_detail)
                row['Description'] = cleaned_detail  # Overwrite the original 'Description' column with cleaned data
                customer_name = extract_customer_name(cleaned_detail, mappings)
                row['Customer Name'] = customer_name if customer_name else 'Not found'
                writer.writerow(row)

        shutil.move(temp_file.name, csv_file)

# Manually created mapping of descriptions to customer names
description_to_customer = {
    '5GUYS': 'Five Guy\'s',
    '7ELEVEN': '7 Eleven',
    '7-ELEVEN': '7 Eleven',
    'ACADEMYSPORTS': 'Academy Sports',
    'ADVANCEAUTOPARTS': 'Advance Auto Parts',
    'AFTERPAY': 'Afterpay',
    'ALDI': 'Aldi',
    'APPLE': 'Apple',
    'AMAZON.COM': 'Amazon',
    'AMAZONPRIME': 'Amazon Prime',
    'AMAZONMUSIC': 'Amazon Music',
    'AMZNFREETIME': 'Amazon Free Time',
    'AMZNMKT': 'Amazon',
    'AMERICANEXPRESS': 'American Express',
    'AMC': 'AMC',
    'APPLE.COM': 'Apple',
    'APPLEPAY': 'Apple Pay',
    'ARBY\'S': 'Arby\'s',
    'BABYLANE\'SCHILDREN\'S': 'Baby Lane\'s Children\'s',
    'BADCOCKHOMEFURNITURE': 'Badcock Home Furniture',
    'BERRYSWEET': 'Berrysweet',
    'BESTBUY': 'Best Buy',
    'BESTWESTERN': 'Best Western',
    'BIGCHIC': 'Big Chic',
    'BLACKBIRD/ROJILIOSME': 'Blackbird/Rojilios Me',
    'BP': 'BP',
    'BUILD-A-BEARWORKSHOP': 'Build A Bear',
    'BURGERKING': 'Burger King',
    'CALHOUNCOFFEECOMPANY': 'Calhoun Coffee Company',
    'CANVA': 'Canva',
    'CAPITALONE': 'Capital One',
    'CHASE': 'Chase',
    'CHECK': 'Manual Check',
    'CHEDDARS': 'Cheddar\'s',
    'CHEVRON': 'Chevron',
    'CHICK-FIL-A': 'Chick-Fil-A',
    'CIRCLEK': 'Circle K',
    'CITGO': 'Citigo',
    'COCACOLA': 'Vending Machine',
    'COSTCO': 'Costco',
    'CREDITONE': 'Credit One',
    'CVS': 'CVS',
    'DAILYDEALZOFMACON': 'Daily Dealz of Macon',
    'DAIRYQUEEN': 'Dairy Queen',
    'DEPOSIT': 'Deposit',
    'DISCOVER': 'Discover',
    'DISNEYPLUS': 'Disney+',
    'DOLLAR-GENERAL': 'Dollar General',
    'DOLLARGENERAL': 'Dollar General',
    'DOLLARTREE': 'Dollar Tree',
    'DOORDASH': 'DoorDash',
    'DUNKIN': 'Dunkin',
    'ELEGANTEXPRESSIONS': 'Elegant Expressions',
    'ETSY': 'Etsy',
    'EXXON': 'Exxon',
    'FACEBOOK': 'Facebook',
    'FANDANGO': 'Fandango',
    'FEDEX': 'FedEx',
    'FOODLION': 'Food Lion',
    'GERBERLIFER': 'Gerber Life Insurance',
    'GERBERLIFER': 'Gerber Life Insurance',
    'GASTTAX': 'GA State Tax',
    'GOOGLEPAY': 'Google Pay',
    'GROUPON': 'Groupon',
    'GREATWALL': 'Great Wall',
    'HAIRTRENDS': 'Hair Trends',
    'HOMETOWNPRINTERS': 'Hometown Printers',
    'HOMEDEPOT': 'Home Depot',
    'HOBBY-LOBBY': 'Hobby Lobby',
    'HULU': 'Hulu',
    'IHOP': 'IHop',
    'IMT': 'Matthew Jina',
    'INGLES': 'Ingles',
    'INMATE': 'Matthew Jina',
    'INSTAGRAM': 'Instagram',
    'JCPENNEY': 'JCPenney',
    'JENNYSCHOPSHOP': 'Jenny\'s Chop Shop',
    'JUMPFORFUN': 'Jump for Fun',
    'JUSTIN\'SPLACE': 'Justin\'s Place',
    'KOHL': 'Kohl\'s',
    'KROGER': 'Kroger',
    'LAFIESTA': 'La Fiesta',
    'LAPARILLA': 'La Parilla',
    'LENDMARKFINANCIAL': 'Lendmark Financial',
    'LENDMARKFINANCIAL': 'Lendmark Financial',
    'LITTLECAESARS': 'Little Caesars',
    'LONGHORN': 'Long Horn Steak House',
    'LOULOU\'SCATERING': 'Loulou\'s Catering',
    'LOWES': 'Lowe\'s',
    'MARATHON': 'Marathon',
    'MARSHALLS': 'Marshalls',
    'MASTERCARD': 'Mastercard',
    'MCDONALD': 'McDonald\'s',
    'MIDLAND': 'Midland Credit',
    'NETFLIX': 'Netflix',
    'NORRIS\'S': 'Norri\'s',
    'O\'CHARLEY\'S': 'O\'Charley\'s',
    'O\'REILLY': 'O\'Reilly\'s',
    'OSAKA': 'Osaka',
    'PARACORDPLANET': 'Paracord Planet',
    'PARLEVEL': 'Vending Machine',
    'PAYPAL': 'PayPal',
    'PEACHTREECAFE': 'Peachtree Café',
    'PERANNA': 'Transfer',
    'PIGGLYWIGGLY': 'Piggy Wiggly',
    'PIIGGIEPARK': 'Piigie Park',
    'PIZZAHUT': 'Pizza Hut',
    'POPSHELF': 'Pop Shelf',
    'PRIMEVIDEO': 'Prime Video',
    'PRIMEVIDEOCHANNELS': 'Prime Video Channels',
    'PROCEEDSOFCLUBACCOUNT': 'Christmas Club',
    'QT': 'Quick Trip',
    'RACEWAY': 'Raceway',
    'ROBINHOOD': 'Robinhood',
    'ROCKAUTO': 'Rock Auto',
    'SABARROS': 'Sabarros',
    'SALTWATERMARKET': 'Salt Water Market',
    'SAMSCLUB': 'Sam\'s Club',
    'SAM\'SCLUB': 'Sam\'s Club',
    'SERVICECHARGE': 'Bank Fee',
    'SHELL': 'Shell',
    'SHEIN': 'Shein',
    'SHUTTERFLY': 'Shutterfly',
    'SLICES': 'Slices',
    'SNAPFINANCE': 'SnapFinance',
    'SPECTRUM': 'Spectrum',
    'SQUARE': 'Square',
    'STAPLES': 'Staples',
    'STARBUCKS': 'Starbucks',
    'STATEFARM': 'State Farm',
    'SUBWAY': 'Subway',
    'TACOBELL': 'Taco Bell',
    'TARGET': 'Target',
    'TASTYSHOPPE': 'Tasty Shoppe',
    'TEMU': 'Temu',
    'THECOUNTRYCUPBOARD': 'The Country Cupboard',
    'THEPEOPLESBANK': 'The People\'s Bank',
    'TIDALWAVE': 'Tidal Wave',
    'TIREMART': 'Tire Mart',
    'TJMAXX': 'T.J. Maxx',
    'TOYSRUS': 'Toys "R" Us',
    'TRANSF': 'Transfer',
    'TRANSPORT': 'Transport',
    'TRANSFTOCLUBSAV': 'Christmas Savings',
    'TRUE-VALUE': 'True Value',
    'TWITTER': 'Twitter',
    'UNITEDBANK': 'Cash',
    'UPS': 'UPS',
    'USPOSTALSERVICE': 'US Postal Service',
    'USPS': 'USPS',
    'VALERO': 'Valero',
    'VENMO': 'Venmo',
    'VISA': 'Visa',
    'WALGREENS': 'Walgreens',
    'WAL-MART': 'Walmart',
    'WALMART': 'Walmart',
    'WMTPLUS': 'Walmart Plus',
    'WAFFLEHOUSE': 'Waffle House',
    'WHATABURGER': 'Whataburger',
    'WENDY\'S': 'Wendys',
    'YODERS': 'Yoders',
    'YOUTUBE': 'YouTube',
    'ZAXBY\'S': 'Zaxby\'s',
    'ZAXBYS': 'Zaxby\'s',
    'ZAZZLE': 'Zazzle',
}

# Replace 'your_file.csv' with the path to your CSV file
csv_file = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Desktop\Transaction Data.csv"

clean_csv_file(csv_file, description_to_customer)
