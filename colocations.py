import os

def remove_paragraph_indents(text):
    cleaned_lines = [line.lstrip() for line in text.splitlines()]
    return 'n'.join(cleaned_lines)

folder_path = r'C:\Users\Lenovo\Documents\корпус персидской поэзии\attar'  
texts = []

for filename in os.listdir(folder_path):
    if filename.endswith('.txt'):
        with open(os.path.join(folder_path, filename), 'r', encoding='utf-8') as file:
            content = file.read()
            cleaned_content = remove_paragraph_indents(content)
            texts.append(cleaned_content)

corpus = ' '.join(texts)
stop_words = set(['آنجا','آنکه','های','ست','کرده','ازین','ازان','باز','گردد','،','گشت','فی','زین','کو','هست','چنین','برای','شود','کس','کز','بهر','همچو','سوی','پی','همی','شمارهٔ','یا','چو','مر','گفت','برات','چون', 'هیچ', 'صد','دارد','داری', 'نیست', 'است','گوید','چی', 'چه', 'اگر', 'هر', 'را', 'از', 'ز', 'چهار', 'سه', 'یک', 'شما', 'وی', 'انها', 'ما', 'او', 'تو', 'من', 'دو', 'پیش', 'گر', 'می', 'و', 'در', 'کی', 'شد', 'کرد', 'گرچه', 'گاه', 'گه', 'زان', 'که', 'به', 'بر', 'آن', 'تا', 'این', 'اندر', 'آندر', 'اینجا', 'درین', 'بود', 'بی', 'آمد', 'چون', 'با', 'ای', 'هم', 'باشد', 'ترا', '-', ':', 'سر'])

def filter_stop_words(text):
    words = text.split()
    return [word for word in words if word not in stop_words]

filtered_words = filter_stop_words(corpus)

from collections import Counter

word_freq = Counter(filtered_words)

def find_context_words(corpus, keyword, window_size=40):
    words = corpus.split()
    context_words = []
    
    for i, word in enumerate(words):
        if word == keyword:
            start_index = max(0, i - window_size)
            end_index = min(len(words), i + window_size + 1)
            context_words.extend(words[start_index:i])
            if i != len(words) - 1:
              context_words.extend(words[i + 1:end_index])
    
    return context_words

keyword = 'قدرت' 
context_words = find_context_words(corpus, keyword)

context_filtered = filter_stop_words(' '.join(context_words))

context_word_freq = Counter(context_filtered)

print("Частота слов в контексте ключевого слова:")
for word, freq in context_word_freq.most_common(15):
    print(f"{word}: {freq}")
