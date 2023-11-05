import os
import openpyxl
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime  # Import datetime module for timestamp
import requests

# get the API KEY here: https://developers.google.com/custom-search/v1/overview
API_KEY = "YOUR API KEY HERE"
# get your Search Engine ID on your CSE control panel
SEARCH_ENGINE_ID = "YOUR SEARCH ENGINE ID HERE"


def perform_google_search(keyword):
    try:

        url = f"https://www.googleapis.com/customsearch/v1?key={API_KEY}&cx={SEARCH_ENGINE_ID}&q={keyword}&start=1"
        results = requests.get(url).json()
        print(results)

        return results
    except Exception as e:
        print("Error", f"Error performing Google search for '{keyword}': {e}")
        messagebox.showerror(
            "Error", f"Error performing Google search for '{keyword}': {e}"
        )
        return []


def extract_info(search_item):
    try:
        long_description = search_item["pagemap"]["metatags"][0]["og:description"]
    except KeyError:
        long_description = "N/A"
    # get the page title
    title = search_item.get("title")
    # page snippet
    snippet = search_item.get("snippet")
    # alternatively, you can get the HTML snippet (bolded keywords)
    html_snippet = search_item.get("htmlSnippet")
    # extract the page url
    link = search_item.get("link")
    return link, title, long_description


def write_to_excel(keyword, results, sheet, header_written=False):
    # Write headers to the sheet if not written before
    if header_written:
        sheet.append(["Keyword", "Rank", "Title", "URL", "Description"])

    # Write the keyword to each row
    for index, result in enumerate(results, start=1):
        url, title, description = extract_info(result)
        sheet.append([keyword, index, title, url, description])


def run_search(keywords, output_folder, wb):

    # Create a new sheet for each search
    sheet = wb.create_sheet(title="Results")
    header_written = True
    for i, keyword in enumerate(keywords):

        print(f"Searching for: {keyword.strip()}")
        results = perform_google_search(keyword.strip())
        # print(results)
        # Pass header_written as an argument to write_to_excel
        search_items = results.get("items")
        write_to_excel(keyword, search_items, sheet, header_written)
        header_written = False

    # Use datetime to get current date and time for the file name
    current_datetime = datetime.now().strftime("%Y%m%d_%H%M%S")
    # output_path = filedialog.asksaveasfilename(initialdir=output_folder, defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile=f"Google_{current_datetime}")
    output_filename = f"Google_{current_datetime}.xlsx"
    output_path = os.path.join(output_folder, output_filename)
    if output_path:
        try:
            wb.save(output_path)
            messagebox.showinfo(
                "Success", "Google search results saved successfully.")
        except Exception as e:
            print("Error", f"Error saving Excel file: {e}")
            messagebox.showerror("Error", f"Error saving Excel file: {e}")


def browse_and_set_variable(var, dialog_method):
    file_path = dialog_method()
    if file_path:
        var.set(file_path)


def create_ui():
    root = tk.Tk()
    root.title("Google Search Results")

    # Create a new Excel workbook
    wb = openpyxl.Workbook()

    file_path_var = tk.StringVar()
    output_folder_var = tk.StringVar()

    label_file = tk.Label(root, text="Select Keywords File:")
    label_file.pack()

    entry_file = tk.Entry(root, textvariable=file_path_var)
    entry_file.pack()

    label_output = tk.Label(root, text="Select Output Folder:")
    label_output.pack()

    entry_output = tk.Entry(root, textvariable=output_folder_var)
    entry_output.pack()

    button_browse = tk.Button(
        root,
        text="Browse Keywords File",
        command=lambda: browse_and_set_variable(
            file_path_var, filedialog.askopenfilename
        ),
    )
    button_browse.pack()

    button_browse_output = tk.Button(
        root,
        text="Browse Output Folder",
        command=lambda: browse_and_set_variable(
            output_folder_var, filedialog.askdirectory
        ),
    )
    button_browse_output.pack()

    button_run = tk.Button(
        root,
        text="Run",
        command=lambda: run_search(
            get_keywords(file_path_var.get()), output_folder_var.get(), wb
        ),
    )
    button_run.pack()

    root.mainloop()


def get_keywords(file_path):
    try:
        with open(file_path, "r") as file:
            return [line.strip() for line in file]
    except Exception as e:
        print("Error", f"Error reading keywords file: {e}")
        messagebox.showerror("Error", f"Error reading keywords file: {e}")
        return []


if __name__ == "__main__":
    create_ui()
