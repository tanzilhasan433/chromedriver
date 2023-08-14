from selenium import webdriver
from selenium.webdriver.common.keys import Keys

import pandas as pd

# Set up WebDriver
driver = webdriver.Chrome()

def get_suggestions(search_query):
    # Open the search engine
    driver.get("https://www.google.com")

    # Enter search query and get suggestions
    search_box = driver.find_element("name", "q")  # Find the search input element by its name attribute

    search_box.send_keys(search_query)

    # Wait for suggestions to load (you might need to adjust the wait time)
    driver.implicitly_wait(5)

    # Get suggestion elements
    suggestion_elements = driver.find_elements(by="css selector", value="ul.G43f7e li")

    suggestions = [element.text for element in suggestion_elements]
    all_suggestion = []

    # Print or process suggestions
    for suggestion in suggestions:
        suggestion_txt = suggestion.split("\n")[0]
        all_suggestion.append(suggestion_txt)
        
    sorted_suggestion = sorted(all_suggestion, key=len)
    shortest = sorted_suggestion[0]
    largest = sorted_suggestion[-1]
    
    return [largest, shortest]


def main():
    excel_file_path = "Excel.xlsx"
    # Open the Excel file
    excel_file = pd.ExcelFile(excel_file_path)
    # List sheet names
    sheet_names = excel_file.sheet_names
    with pd.ExcelWriter(excel_file_path) as writer:
        for sheet_name in sheet_names:
            data = excel_file.parse(sheet_name)

            third_column = data.iloc[1:, 2]
            
            for i,each in third_column.items():
                result = get_suggestions(each)
                data.iloc[i, 3] = result[0]
                data.iloc[i, 4] = result[1]

            print(data)
            data.columns = [''] * len(data.columns)       
            data.to_excel(writer, sheet_name, index=False)
        writer.close()
    excel_file.close()
             
if __name__ == '__main__':
    main()
    # Close WebDriver
    driver.quit()
