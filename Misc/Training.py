import subprocess
import sys

required_packages = ['nltk', 'contractions']

def is_package_installed(package):
    try:
        subprocess.check_output([sys.executable, '-m', 'pip','show', package])
        return True
    except subprocess.CalledProcessError:
        return False
    
for package in required_packages:
    if not is_package_installed(package):
        #print(f"Installing {package}...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])

import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
import string
import pandas as pd
import contractions

df = pd.read_csv(r'C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Desktop\Conversation.csv')

def clean(text):
    text = text.lower()
    text = contractions.fix(text)
    text = ''.join([char for char in text if char not in string.punctuation])
    return text

def tokenize(text):
    # Tokenization
    tokens = word_tokenize(text)

    # Removing stopwords (common words like "the", "and", "is" that may not be relevant)
    stop_words = set(stopwords.words('english'))
    filtered_tokens = [word for word in tokens if word.lower() not in stop_words]

    # Removing punctuation
    filtered_tokens = [word for word in filtered_tokens if word not in string.punctuation]

    cleaned_text =''.join(filtered_tokens)

    return cleaned_text

df['question'] = df['question'].apply(clean)
df['answer'] = df['answer'].apply(clean)
df['question'] = df['question'].apply(tokenize)
df['answer'] = df['answer'].apply(tokenize)

df.to_csv(r'cleaned_text.csv', index=False)