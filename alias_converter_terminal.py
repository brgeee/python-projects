import re
from spellchecker import SpellChecker
import json

CUSTOM_DICT_FILE = "C:/Users/EB1801445/OneDrive - Oracle Corporation/Documents/Documents/Useful Python Scripts/custom_dictionary.json"

def load_custom_dictionary():
    try:
        with open(CUSTOM_DICT_FILE, 'r') as f:
            return set(json.load(f))
    except FileNotFoundError:
        return set()
    
def save_custom_dictionary(custom_dict):
    with open(CUSTOM_DICT_FILE, 'w') as f:
        json.dump(list(custom_dict), f)

def match_case(word, corrected_word):
    if word.isupper():
        return corrected_word.upper()
    elif word.islower():
        return corrected_word.lower()
    elif word.istitle():
        return corrected_word.capitalize()
    else:
        return corrected_word

def convert_name(name, append_text, use_spellcheck=True):
    spell = SpellChecker()

    custom_dict = load_custom_dictionary()

    #add words to a dictionary
    custom_words = ['IgM', 'IgG', 'IgD', 'IgE', 'IgA', 'CSF', 'rRNA', 'DNA'] + list(custom_dict)
    for word in custom_words:
        spell.word_frequency.add(word.lower())

    #handle the conversion of '%' only for the alias
    alias_name = name.replace('%', ' PCT')

    #remove any text within parentheses only for alias_name
    alias_name = re.sub(r'\(.*?\)', '', alias_name)

    #replace various characters with _
    alias_name = re.sub(r'[/,\.\-\+]', '_', alias_name)

    #split the name and alias_name by spaces
    words = name.split()
    alias_words = alias_name.split()

    #process each word with or without spellcheck
    corrected_words = []
    final_corrected_words = []
    alias_corrected_words = []
    any_corrections = False
    
    for i in range(len(words)):
        original_word = words[i]
        alias_word = alias_words[i] if i < len(alias_words) else words[i]

        if original_word in custom_words:
            corrected_words.append(original_word)
            final_corrected_words.append(original_word)
            alias_corrected_words.append(original_word.upper())
        else:
            if use_spellcheck:
                corrected_word = spell.correction(original_word)

                #if no suggestion or the same word is returned, skip to the next word
                if corrected_word is None or corrected_word == original_word:
                    corrected_words.append(original_word)
                    final_corrected_words.append(original_word)
                    alias_corrected_words.append(alias_word.upper())
                    continue
                
                if original_word != corrected_word and not re.search(r'\d', original_word):
                    any_corrections = True
                    print(f"Potential misspelling detected: '{original_word}'")
                    response = input(f"Would you like to update it to '{corrected_word}'? (y/n): ")
                    if response.lower() == 'y':
                        corrected_word = match_case(original_word, corrected_word)
                        original_word = corrected_word
                        alias_word = corrected_word
                    else:
                        add_to_dict = input(f"Do you want to add '{original_word}' to the dictionary? (y/n): ")
                        if add_to_dict.lower() == 'y':
                            custom_dict.add(original_word.lower())
                            spell.word_frequency.add(original_word.lower())
                            save_custom_dictionary(custom_dict)
                            print(f"Adding '{original_word}' to the custom dictionary.")
                            if original_word.lower() in spell:
                                print(f"'{original_word}' successfully added to the dictionary.")
                            else:
                                print(f"Failed to add '{original_word}' to the dictionary.")
            else:
                corrected_word = original_word

            #handle camel case words for alias, and convert to upper case
            if re.search(r'[a-z][A-Z]', original_word):
                alias_word = re.sub(r'([a-z])([A-Z])', r'\1_\2', alias_word).upper()
            else:
                alias_word = alias_word.upper()

            corrected_words.append(original_word)
            final_corrected_words.append(original_word)
            alias_corrected_words.append(alias_word)

    #join words with underscores
    processed_name = '_'.join(corrected_words)
    processed_alias = '_'.join(alias_corrected_words)

    #append text to the end
    processed_alias = f"{processed_alias}_{append_text.upper()}"

    original_name_corrected = " ".join(final_corrected_words)

    #character warning if >100 chars in name/alias
    if len(original_name_corrected) > 100:
        print(f"Warning: Concept Name is {len(original_name_corrected)} characters.")

    if len(processed_alias) > 100:
        print(f"Warning: Concept Alias is {len(processed_alias)} characters.")

    return original_name_corrected, processed_alias

def main():
    use_spellcheck = input("Do you want to use spell check? (y/n): ").strip().lower() == 'y'
    append_text = input("Enter the text you want to append (e.g., OBSTYPE): ").strip()
    
    while True:
        name = input("Enter the name you want to convert (or type 'stop' to exit): \n")
        
        if name.strip().lower() == 'stop':
            print("Exiting the program.")
            break

        converted_name, converted_alias = convert_name(name, append_text, use_spellcheck)
        print(f"\nConverted Name:\n{converted_name}")
        print(f"\nConverted Alias:\n{converted_alias}\n")

if __name__ == "__main__":
    main()
