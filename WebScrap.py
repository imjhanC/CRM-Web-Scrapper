from tkinter import Tk, ttk
from tkinter import messagebox
import tkinter as tk
import undetected_chromedriver as uc 
import time
import threading 
import os
import sys
from tkinter import *
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.common.action_chains import ActionChains 
from selenium.common.exceptions import TimeoutException, NoSuchElementException , StaleElementReferenceException

# FUNCTIONS TODO
# This is the windows for editing products that are currently on the CRM
def edit_prod_crm(row_data,products_info,driver=None,update_treeview=None):
    # This is a recursive call for this main function 
    def edit_comment_name():
        def search_comment_item(event):
            search_query = search_comment_entry.get().lower()
            comment_listbox.delete(0, END)

            if search_query:
                with open("Filtered_data.txt", "r") as file:
                    found = False
                    for line in file:
                        item_name = line.split(">")[0]  # Get only the first column
                        if search_query in item_name.lower():
                            comment_listbox.insert(END, item_name)
                            found = True
                    if not found:
                        comment_listbox.insert(END, "No results found")
            else:
                comment_listbox.insert(END, "No results found")

        def select_comment_items():
            selected_items = comment_listbox.curselection()
            for item_index in selected_items:
                item_name = comment_listbox.get(item_index)
                if item_name != "No results found":
                    # Check if there is already content in the text_comment widget
                    current_content = text_comment.get("1.0", END).strip()
                    if current_content:
                        # If there is content, add a newline before the new item
                        text_comment.insert(END, "\n" + item_name)
                    else:
                        # If no content, simply add the item without a newline
                        text_comment.insert(END, item_name)
            root_comment.destroy()  # Close the edit window

        def cancel_comment_edit():
            root_comment.destroy()

        # Create the edit comment window
        root_comment = Toplevel(root)
        root_comment.title("Add / Delete Comment Items")
        window_width_comment = 550
        window_height_comment = 400

        screen_width_comment = root_comment.winfo_screenwidth()
        screen_height_comment = root_comment.winfo_screenheight() 

        x_position = (screen_width_comment - window_width_comment) // 2
        y_position = (screen_height_comment - window_height_comment) // 2
        root_comment.geometry(f"{window_width_comment}x{window_height_comment}+{x_position}+{y_position}")
        root_comment.resizable(False, False)
        root_comment.transient(root)
        root_comment.grab_set()

        # Search bar
        Label(root_comment, text="Search Product items", font="BOLD").pack(pady=10)
        search_comment_entry = Entry(root_comment, width=60)
        search_comment_entry.pack(pady=5)
        search_comment_entry.bind("<KeyRelease>", search_comment_item)

        # Listbox for displaying search results
        comment_listbox = Listbox(root_comment, width=60, height=10, selectmode=MULTIPLE)
        comment_listbox.pack(pady=5)

        # Frame for Select and Cancel buttons
        button_frame = Frame(root_comment)
        button_frame.pack(pady=10)

        # Select and Cancel buttons
        select_button = Button(button_frame, text="Select", command=select_comment_items)
        select_button.grid(row=0, column=0, padx=5)

        cancel_button = Button(button_frame, text="Cancel", command=cancel_comment_edit)
        cancel_button.grid(row=0, column=1, padx=5)
    
    def search_item(event):
        search_query = search_entry.get().lower()
        suggestions_listbox.delete(0, END)

        if search_query:  # Proceed only if there is a search query
            with open("Filtered_data.txt", "r") as file:
                for line in file:
                    item_name = line.split(">")[0]  # Get only the first column
                    if search_query in item_name.lower():
                        suggestions_listbox.insert(END, item_name)  # Add to Listbox if query matches
        else:
            suggestions_listbox.insert(END, "No results found")  # Show if no matching items
    
    # Function to handle item selection from the Listbox
    def select_item(event):
        try:
            selected_item = suggestions_listbox.get(suggestions_listbox.curselection())
            search_entry.delete(0, END)
            search_entry.insert(0, selected_item)  # Display the selected item in the search bar
        except TclError:
            pass  # Ignore if there is no selection
    
    # Function to open the item edit window
    def edit_item_name():
        def cancel_edit():
            root_edit.destroy()  # Close the edit window

        def select_and_close():
            try:
                selected_item = suggestions_listbox.get(suggestions_listbox.curselection())
                entry_item_name.config(text=selected_item)  # Set the selected item to the main label
                root_edit.destroy()  # Close the edit window
            except TclError:
                messagebox.showerror("Selection Error", "Please select a valid item from the list.",parent=root_edit)

        root_edit = Toplevel(root)
        root_edit.title("Select new item's name")
        window_width_edit = 550
        window_height_edit = 400

        screen_width_edit = root_edit.winfo_screenwidth()
        screen_height_edit = root_edit.winfo_screenheight() 

        x_position = (screen_width_edit - window_width_edit) // 2
        y_position = (screen_height_edit - window_height_edit) // 2
        root_edit.geometry(f"{window_width_edit}x{window_height_edit}+{x_position}+{y_position}")
        root_edit.resizable(False, False)

        root_edit.transient(root)  # Keep the edit window on top of the main window
        root_edit.grab_set()       # Disable interaction with the main window while edit window is open

        # Search bar
        Label(root_edit, text="Search Product's name", font="BOLD").pack(pady=10)
        global search_entry  # Make search_entry global to access in search_item function
        search_entry = Entry(root_edit, width=60)
        search_entry.pack(pady=5)
        search_entry.bind("<KeyRelease>", search_item)  # Bind key release event for dynamic search

        # Listbox for suggestions
        global suggestions_listbox
        suggestions_listbox = Listbox(root_edit, width=60, height=10)
        suggestions_listbox.pack(pady=5)
        suggestions_listbox.bind("<<ListboxSelect>>", select_item)  # Bind selection event

        # Frame for "Select" and "Cancel" buttons
        button_frame = Frame(root_edit)
        button_frame.pack(pady=10)

        # "Select" button
        select_button = Button(button_frame, text="Select", command=select_and_close)
        select_button.grid(row=0, column=0, padx=5)

        # "Cancel" button
        cancel_button = Button(button_frame, text="Cancel", command=cancel_edit)
        cancel_button.grid(row=0, column=1, padx=5)
    
    # This function is to clear the comment's text 
    def clear_comment():
        text_comment.delete("1.0", END)  # Clear all text in the text_comment widget

    def extract_products_info_edit_windows():
        """
        Extract product information for all rows in the table with explicit waits.
        Added additional waits and checks for recently edited rows.
        """
        wait = WebDriverWait(driver, 20)
        products = []
        
        try:
            # Wait for initial load and line items
            print("Waiting for line items to load...")
            wait.until(EC.presence_of_element_located((By.CLASS_NAME, "lineItemRow")))
            #time.sleep(2)  # Short pause to ensure full load
            
            # Find all rows
            rows = driver.find_elements(By.CLASS_NAME, "lineItemRow")
            total_rows = len(rows)
            print(f"Found {total_rows} line items to process")
            
            for index, row_element in enumerate(rows, 1):
                row_number = int(row_element.get_attribute('id').replace('lineItemRow', ''))
                print(f"Processing product {row_number} ({index}/{total_rows})")
                
                # Special scroll handling for last row
                if index == total_rows:
                    driver.execute_script("arguments[0].scrollIntoView(false);", row_element)
                else:
                    driver.execute_script("arguments[0].scrollIntoView(true);", row_element)
                
                # Add pause after scrolling
                #time.sleep(1)
                
                # Extract product information
                try:
                    product_info = {}
                    base_id = f"lineItemRow{row_number}"
                    
                    # Wait for row to be present after scroll
                    wait.until(EC.presence_of_element_located((By.ID, base_id)))
                    
                    # Extract basic information first
                    quantity_element = wait.until(EC.presence_of_element_located((
                        By.XPATH, f"//tr[@id='{base_id}']//input[@name='qty{row_number}']"
                    )))
                    product_info['quantity'] = quantity_element.get_attribute('value')
                    
                    item_name_element = wait.until(EC.presence_of_element_located((
                        By.XPATH, f"//tr[@id='{base_id}']//input[@readonly='readonly' and contains(@class, 'form-control bg-white')]"
                    )))
                    product_info['item_name'] = item_name_element.get_attribute('title')
                    
                    # Extract subproducts if present
                    try:
                        subproducts_element = driver.find_element(By.CSS_SELECTOR, f"#lineItemRow{row_number} .subProducts")
                        if subproducts_element.is_displayed():
                            subproducts = []
                            for subproduct_div in subproducts_element.find_elements(By.CSS_SELECTOR, 'div'):
                                subproduct_text = subproduct_div.text.strip()
                                if subproduct_text:
                                    subproducts.append(subproduct_text)
                            product_info['subproducts'] = subproducts
                    except NoSuchElementException:
                        product_info['subproducts'] = []
                    
                    # Extract comment
                    comment_element = wait.until(EC.presence_of_element_located((
                        By.XPATH, f"//textarea[@name='comment{row_number}']"
                    )))
                    product_info['comment'] = comment_element.get_attribute('value')
                    
                    # For price-related elements, add extra wait and retry logic
                    max_retries = 3
                    retry_delay = 1  # seconds
                    
                    for attempt in range(max_retries):
                        try:
                            # Wait for price calculations to complete
                            price_element = wait.until(EC.presence_of_element_located((
                                By.XPATH, f"//tr[@id='{base_id}']//input[@test-attr='listPrice{row_number}']"
                            )))
                            
                            # Force a refresh of price calculations if needed
                            driver.execute_script("""
                                var element = arguments[0];
                                element.dispatchEvent(new Event('change', { 'bubbles': true }));
                            """, price_element)
                            
                            time.sleep(retry_delay)  # Wait for calculations
                            
                            # Try to get all price-related values
                            product_info['unit_selling_price'] = price_element.get_attribute('test-attr-price')
                            
                            total_cost_element = wait.until(EC.presence_of_element_located((
                                By.XPATH, f"//tr[@id='{base_id}']//input[@type='number' and @name='purchase_cost{row_number}']"
                            )))
                            product_info['total_purchase_cost'] = total_cost_element.get_attribute('value')
                            
                            unit_cost_element = wait.until(EC.presence_of_element_located((
                                By.XPATH, f"//tr[@id='{base_id}']//input[@name='unit_purchase_cost{row_number}']"
                            )))
                            product_info['unit_purchase_cost'] = unit_cost_element.get_attribute('value')
                            
                            # Wait for calculated fields with explicit wait for value changes
                            discount_script = f"return document.getElementById('totalAfterDiscount{row_number}').textContent.trim()"
                            total_discount_element = WebDriverWait(driver, 5).until(
                                lambda d: d.execute_script(discount_script) != ""
                            )
                            product_info['total_after_discount'] = driver.execute_script(discount_script).strip()
                            
                            margin_script = f"return document.getElementById('margin{row_number}').textContent.trim()"
                            margin_element = WebDriverWait(driver, 5).until(
                                lambda d: d.execute_script(margin_script) != ""
                            )
                            product_info['margin'] = driver.execute_script(margin_script).strip()
                            
                            tax_element = wait.until(EC.presence_of_element_located((
                                By.XPATH, f"//td[@id='taxTotal{row_number}']//button"
                            )))
                            product_info['tax'] = tax_element.text
                            
                            net_price_script = f"return document.getElementById('netPrice{row_number}').textContent"
                            net_price_element = WebDriverWait(driver, 5).until(
                                lambda d: d.execute_script(net_price_script) != ""
                            )
                            product_info['net_price'] = driver.execute_script(net_price_script)
                            
                            break  # If we got here, we successfully got all values
                            
                        except (TimeoutException, StaleElementReferenceException) as e:
                            if attempt == max_retries - 1:  # Last attempt
                                print(f"Failed to get price information for row {row_number} after {max_retries} attempts: {str(e)}")
                                # Set default values or handle the error as needed
                                product_info.update({
                                    'unit_selling_price': '0',
                                    'total_after_discount': '0',
                                    'margin': '0',
                                    'tax': '0',
                                    'net_price': '0'
                                })
                            else:
                                print(f"Retry {attempt + 1} for row {row_number}")
                                time.sleep(retry_delay)
                                continue
                    
                    products.append(product_info)
                    print(f"Successfully extracted information for product {row_number}")
                    
                except (TimeoutException, NoSuchElementException) as e:
                    print(f"Error extracting data for row {row_number}: {str(e)}")
                    continue
                    
            return products
            
        except TimeoutException as e:
            print(f"Timeout waiting for line items: {str(e)}")
            return []
        except Exception as e:
            print(f"Error in extract_products_info: {str(e)}")
            return []

    def on_confirm_click():
        # Get the edited quantity from the Spinbox widget
        edited_quantity = spinbox_quantity.get()
        item_name = entry_item_name.cget("text")  # Get text from the Label widget
        comment = text_comment.get("1.0", "end").strip()  # Get text from Text widget and strip extra spaces
        unit_selling_price = entry_unit_selling_price.get()
        #total_purchase_cost = entry_total_purchase_cost.get()
        #unit_purchase_cost = entry_unit_purchase_cost.get()

        # Fetch the row_number from label_no_value
        row_number_no = label_no_value.cget("text")  # Get text from label_no_value, which corresponds to row_data[0]

        # Define base_id using row_number
        base_id = f"lineItemRow{row_number_no}"
        
        # Ensure product_info dictionary exists
        product_info = {}
        product_info['item_name'] = item_name
        product_info['comment'] = comment
        product_info['unit_selling_price'] = unit_selling_price
        #product_info['total_purchase_cost'] = total_purchase_cost
        #product_info['unit_purchase_cost'] = unit_purchase_cost
        product_info['quantity'] = edited_quantity
        wait = WebDriverWait(driver,20)

        try:
            # Wait for the row to be present on the web page
            wait.until(EC.presence_of_element_located((By.ID, base_id)))
            row_element = wait.until(EC.presence_of_element_located((By.ID, base_id)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", row_element)

            # Step 1: Locate the item name element (readonly text field)
            item_name_element = wait.until(EC.presence_of_element_located((
                By.XPATH, 
                f"//tr[@id='{base_id}']//input[@readonly='readonly' and contains(@class, 'form-control bg-white')]"
            )))
            
            # Get the current item name from the element's 'title' attribute
            current_item_name = item_name_element.get_attribute('title')

            # Check if the current item name matches the desired item_name
            if current_item_name != item_name:
                # Proceed with item name replacement only if different

                # Clear the readonly field by clicking the "fa-times-circle" icon
                clear_icon = wait.until(EC.element_to_be_clickable((
                    By.XPATH, 
                    f"//tr[@id='{base_id}']//i[contains(@class, 'fa-times-circle')]"
                )))
                clear_icon.click()

                # Step 3: Click on the dropdown field to activate the search input
                dropdown_field = wait.until(EC.element_to_be_clickable((
                    By.XPATH, 
                    f"//tr[@id='{base_id}']//span[contains(@class, 'select2-selection__rendered')]"
                )))
                dropdown_field.click()

                # Step 4: Type the item name into the search input field and press Enter
                search_input = wait.until(EC.presence_of_element_located((
                    By.XPATH, 
                    f"//tr[@id='{base_id}']//input[@class='select2-search__field']"
                )))
                search_input.send_keys(item_name)
                # click the result in the dropdown menu after keyed in the item name
                first_result = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'select2-results__option--highlighted')]")))
                first_result.click()

            # Update the quantity on the page
            quantity_element = wait.until(EC.presence_of_element_located((
                By.XPATH, 
                f"//tr[@id='{base_id}']//input[@name='qty{row_number_no}']"
            )))

            # Insert Comment
            comment_element = wait.until(EC.presence_of_element_located((
                By.XPATH, f"//textarea[@name='comment{row_number_no}']"
            )))

            # Get the current comment value from the webpage
            current_comment = comment_element.get_attribute('value')
            product_info['comment'] = current_comment  # Store the current value in product_info

            # Check if the current comment is different from the desired comment
            if current_comment != comment:
                # Update the comment only if it's different
                comment_element.clear()  # Clear existing content
                comment_element.send_keys(comment)  # Set new comment

            # Wait for the price element and get the current unit selling price
            price_element = wait.until(EC.presence_of_element_located((
                By.XPATH, f"//tr[@id='{base_id}']//input[@test-attr='listPrice{row_number_no}']"
            )))
            current_unit_selling_price = price_element.get_attribute('test-attr-price')
            product_info['unit_selling_price'] = current_unit_selling_price  # Store the current value

            # Check if the current unit selling price is different from the desired value
            if current_unit_selling_price != unit_selling_price:
                price_element.clear()  # Clear existing content
                price_element.send_keys(unit_selling_price)  # Update with new value


            # Wait for the total purchase cost element and get the current value
            #total_cost_element = wait.until(EC.presence_of_element_located((
            #    By.XPATH, f"//tr[@id='{base_id}']//input[@type='number' and @name='purchase_cost{row_number_no}']"
            #)))
            #current_total_purchase_cost = total_cost_element.get_attribute('value')
            #product_info['total_purchase_cost'] = current_total_purchase_cost

            # Check if the current total purchase cost is different from the desired value
            #if current_total_purchase_cost != total_purchase_cost:
            #    total_cost_element.clear()
            #    total_cost_element.send_keys(total_purchase_cost)

            # Wait for the unit purchase cost element and get the current value
            #unit_cost_element = wait.until(EC.presence_of_element_located((
            #    By.XPATH, f"//tr[@id='{base_id}']//input[@name='unit_purchase_cost{row_number_no}']"
            #)))
            #current_unit_purchase_cost = unit_cost_element.get_attribute('value')
            #product_info['unit_purchase_cost'] = current_unit_purchase_cost

            # Check if the current unit purchase cost is different from the desired value
            #if current_unit_purchase_cost != unit_purchase_cost:
            #    unit_cost_element.clear()
            #    unit_cost_element.send_keys(unit_purchase_cost)

            driver.execute_script("arguments[0].value = arguments[1];", quantity_element, edited_quantity)

            # Optional: Trigger the change event for JavaScript dependencies, if needed
            driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", quantity_element)

            section_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((
                By.CSS_SELECTOR, 
                'input[type="text"][placeholder="Section Name"].form-control.mr-2'
                ))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", section_input)
            #time.sleep(1)
            driver.execute_script("arguments[0].click();", section_input)
            #time.sleep(1)
            fetched_edited_details = extract_products_info_edit_windows()
            if update_treeview:
                update_treeview(fetched_edited_details)
            
            # For debugging 
            print(">>>>>>>>>>>>>>>>>>>>")
            print(fetched_edited_details)

            root.destroy()
        except TimeoutException:
            print("Failed to locate the row or quantity input field.")
        


    root = Tk()
    root.title("Selected products' details")
    root.attributes('-topmost', True)
    window_width = 640
    window_height = 620

    # Get the screen width and height
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight() 

    # Calculate the position to center the window
    x_position = (screen_width - window_width) // 2
    y_position = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
    root.resizable(False, False)

    # Set the main frame for the form
    edit_main_frame = Frame(root)
    edit_main_frame.grid(padx=10, pady=10)

    # No
    label_no = Label(edit_main_frame, text="No")
    label_no.grid(row=0, column=0, sticky=E, padx=5, pady=5)
    label_no_value = Label(edit_main_frame, text=row_data[0], font="BOLD")  # Changed Entry to Label
    label_no_value.grid(row=0, column=1, sticky=W, padx=5, pady=5)

    # Quantity (using Spinbox for increment and decrement)
    label_quantity = Label(edit_main_frame, text="Quantity")
    label_quantity.grid(row=1, column=0, sticky=E, padx=5, pady=5)
    spinbox_quantity = Spinbox(edit_main_frame, from_=1, to=99999, width=10)
    spinbox_quantity.grid(row=1, column=1, sticky=W, padx=5, pady=5)
    spinbox_quantity.delete(0,"end")
    spinbox_quantity.insert(0, row_data[1])
    # Item Name
    label_item_name = Label(edit_main_frame, text="Item Name")
    label_item_name.grid(row=2, column=0, sticky=E, padx=5, pady=20)
    entry_item_name = Label(edit_main_frame, text=row_data[2], font="BOLD",wraplength=250)
    entry_item_name.grid(row=2, column=1, sticky=W, padx=5, pady=20)

    # Add edit button next to Item Name
    button_item_name = Button(edit_main_frame, text="EDIT", width=5, font="BOLD", command=edit_item_name,justify=LEFT)
    button_item_name.grid(row=2, column=3, sticky=W, padx=5)

    formatted_subproducts = "\n".join([f"{subproduct}" for subproduct in row_data[3]])

    # Subproducts
    label_subproducts = Label(edit_main_frame, text="Subproducts")
    label_subproducts.grid(row=3, column=0, sticky=E, padx=5, pady=5)
    text_subproducts = Text(edit_main_frame, width=30, height=4, wrap="word")
    text_subproducts.grid(row=3, column=1, sticky=W, padx=5, pady=5)
    text_subproducts.insert("1.0", formatted_subproducts)
    text_subproducts.config(state="disabled")
    scrollbar_subprod = Scrollbar(edit_main_frame, command=text_subproducts.yview)
    scrollbar_subprod.grid(row=3, column=2, sticky="ns")
    text_subproducts.config(yscrollcommand=scrollbar_subprod.set)

    # Comment
    label_comment = Label(edit_main_frame, text="Comment")
    label_comment.grid(row=4, column=0, sticky=E, padx=5, pady=2)

    # Use Text widget for multiline input
    text_comment = Text(edit_main_frame, width=30, height=4, wrap="word")
    text_comment.grid(row=4, column=1, sticky=W, padx=5, pady=2)
    text_comment.insert("1.0", row_data[4])

    # Set up a scroll bar for the Text widget
    scrollbar_comment = Scrollbar(edit_main_frame, command=text_comment.yview)
    scrollbar_comment.grid(row=4, column=2, sticky="ns")
    text_comment.config(yscrollcommand=scrollbar_comment.set)

    # Add folder button next to Comment scrollbar
    button_comment = Button(edit_main_frame, text="+", width=2, font="BOLD",command=edit_comment_name)
    button_comment.grid(row=4, column=3, sticky=E, padx=10)

    trash_comment = Button(edit_main_frame, text="🗑", width=2, font="BOLD", command=clear_comment)
    trash_comment.grid(row=4, column=4, sticky=W, padx=0)

    # Unit Selling Price
    label_unit_selling_price = Label(edit_main_frame, text="Unit Selling Price")
    label_unit_selling_price.grid(row=5, column=0, sticky=E, padx=5, pady=2)
    entry_unit_selling_price = Entry(edit_main_frame, width=30)
    entry_unit_selling_price.grid(row=5, column=1, sticky=W, padx=5, pady=2)
    entry_unit_selling_price.insert(0 , row_data[5])

    # Item Total Purchase Cost
    #label_total_purchase_cost = Label(edit_main_frame, text="Item Total Purchase Cost")
    #label_total_purchase_cost.grid(row=6, column=0, sticky=E, padx=5, pady=2)
    #entry_total_purchase_cost = Entry(edit_main_frame, width=30)
    #entry_total_purchase_cost.grid(row=6, column=1, sticky=W, padx=5, pady=2)
    #entry_total_purchase_cost.insert(0, row_data[6])

    # Item Unit Purchase Cost
    #label_unit_purchase_cost = Label(edit_main_frame, text="Item Unit Purchase Cost")
    #label_unit_purchase_cost.grid(row=7, column=0, sticky=E, padx=5, pady=2)
    #entry_unit_purchase_cost = Entry(edit_main_frame, width=30)
    #entry_unit_purchase_cost.grid(row=7, column=1, sticky=W, padx=5, pady=2)
    #entry_unit_purchase_cost.insert(0,row_data[7])

    # Total
    label_total = Label(edit_main_frame, text="Total")
    label_total.grid(row=6, column=0, sticky=E, padx=5, pady=2)
    label_total_value = Label(edit_main_frame, text=row_data[8], font="BOLD")  # Changed Entry to Label
    label_total_value.grid(row=6, column=1, sticky=W, padx=5, pady=2)

    # Margin
    label_margin = Label(edit_main_frame, text="Margin")
    label_margin.grid(row=7, column=0, sticky=E, padx=5, pady=2)
    label_margin_value = Label(edit_main_frame, text=row_data[9], font="BOLD")  # Changed Entry to Label
    label_margin_value.grid(row=7, column=1, sticky=W, padx=5, pady=2)

    # Tax
    label_tax = Label(edit_main_frame, text="Tax")
    label_tax.grid(row=8, column=0, sticky=E, padx=5, pady=2)
    label_tax_value = Label(edit_main_frame, text=row_data[10], font="BOLD")  # Changed Entry to Label
    label_tax_value.grid(row=8, column=1, sticky=W, padx=5, pady=2)

    # Net Price
    label_net_price = Label(edit_main_frame, text="Net Price")
    label_net_price.grid(row=9, column=0, sticky=E, padx=5, pady=2)
    label_net_price_value = Label(edit_main_frame, text=row_data[11], font="BOLD")  # Changed Entry to Label
    label_net_price_value.grid(row=9, column=1, sticky=W, padx=5, pady=2)

    # Add a frame for the Cancel and Confirm buttons at the bottom
    button_frame = Frame(root)
    button_frame.grid(row=10, column=0, sticky=E, padx=0, pady=10)

    # Cancel and Confirm buttons
    button_cancel = Button(button_frame, text="Cancel", width=10, command=root.destroy)
    button_cancel.grid(row=0, column=0, padx=5)

    button_confirm = Button(button_frame, text="Confirm", width=10,command=on_confirm_click)
    button_confirm.grid(row=0, column=1, padx=5)

    edit_main_frame.grid_columnconfigure(1,weight =1)
    # Run the main event loop
    root.mainloop()

def save_and_close(driver):
    wait = WebDriverWait(driver, 10)  # Set an explicit wait time of 10 seconds

    try:
        # Part 1: Find and Scroll
        element_to_scroll = wait.until(
            EC.presence_of_element_located((By.CLASS_NAME, "d-flex.align-items-center.justify-content-end.py-12px.px-4"))
        )
        ActionChains(driver).move_to_element(element_to_scroll).perform()
        print("Scrolled to container")
        
        # Part 2: Click "Save" Button
        save_button = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#\\31 266920 > div.bg-white.vds-bodyfixedFooter.w-100 > div > div > button.btn.btn-primary.py-7-5.px-3"))
        )
        
        # Ensure the "Save" button is within the visible portion of the screen
        #driver.execute_script("arg uments[0].scrollIntoView(true);", save_button)  # Scroll button into view
        time.sleep(1)  # Small delay to ensure visibility
        
        # Click the button using JavaScript if normal click doesn't work
        driver.execute_script("arguments[0].click();", save_button)
        print("Clicked 'Save' button")
        
        # Part 2.5: Click "Cancel" Button in the Modal Footer (after Save)
        #cancel_button = wait.until(
        #    EC.element_to_be_clickable((By.CSS_SELECTOR, "footer.modal-footer button.btn.btn-secondary"))
        #)
        #cancel_button.click()
        #print("Clicked 'Cancel' button in modal footer")

        # Wait for the modal to disappear before continuing
        try:
            wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.modal.show")))
            print("Modal closed, continuing with the next steps.")
        except TimeoutException:
            print("Modal did not close in time, continuing with the next steps.")

        # Part 3: Find and Click "Back" Icon
        back_icon = wait.until(
            EC.element_to_be_clickable((By.ID, "previewBack"))
        )
        back_icon.click()
        
        # Part 4: Find and Click Empty Span
        empty_span = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//span[@class='text-center my-3 font-15']"))
        )
        empty_span.click()
        
        # Part 5: Find and Click Logout
        logout = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//div[@title='Logout' and contains(@class, 'name')]"))
        )
        logout.click()
        print("Clicked 'Logout'")

    except Exception as e:
        print("An error occurred:", e)

# This is the windows for managing the products in SOLIDCAM products like updating , adding ,or delete the products 
def select_prod_windows():
    # Function to handle search and highlight matching rows
    def search_and_highlight():
        query = search_var.get().lower()  # Get the search query and convert to lowercase for case-insensitive search
        for item in treeview.get_children():
            treeview.item(item, tags=())  # Remove any previous highlight tags

        # Highlight matching items
        for item in treeview.get_children():
            values = treeview.item(item, 'values')
            if any(query in str(value).lower() for value in values):  # Check if the query matches any value
                treeview.item(item, tags=('highlight',))  # Add the 'highlight' tag to matching rows

        update_item_count()  # Update item count if needed
    
    # Function to clear search and reset highlights
    def clear_search_and_highlight():
        search_var.set("")  # Clear the search entry
        for item in treeview.get_children():
            treeview.item(item, tags=())  # Remove any highlight tags
        update_item_count()

    def update_item_count():
        count = len(treeview.get_children())
        labelItem.config(text=f"Number of items: {count}")
    
    def add_product():
        add_window = Toplevel(root)  # Create a new window for adding products
        add_window.title('Add Product')
        add_window.geometry("500x400")
        add_window.resizable(False, False)
        add_window.grab_set()

        add_frame = ttk.Frame(add_window, padding=10)
        add_frame.pack(fill=BOTH, expand=True)

        # Create entry fields for each attribute
        labels = ["Product Name", "Unit Price", "Purchase Cost", "Product Print Name", "Product Print Descriptions"]
        variables = []
        for i, label in enumerate(labels):
            ttk.Label(add_frame, text=f"{label}: ").grid(row=i, column=0, sticky="e", padx=5, pady=5)
            var = StringVar()
            entry = ttk.Entry(add_frame, textvariable=var, width=40)
            entry.grid(row=i, column=1, sticky="w", padx=5, pady=5)
            variables.append(var)

        def save_product():
            values = [var.get().strip() for var in variables]
            if any(not v for v in values):
                messagebox.showerror("Invalid Input", "All fields are required.")
                return

            # Save to Filtered_data.txt
            try:
                with open('Filtered_data.txt', 'a') as file:
                    file.write('>'.join(values) + "\n")

                # Insert the new product into the Treeview
                treeview.insert("", "end", values=values)
                update_item_count()
                messagebox.showinfo("Add Product", "Product added successfully!")
                add_window.destroy()
            except Exception as e:
                messagebox.showerror("File Error", f"An error occurred while saving the product: {e}")

        save_button = ttk.Button(add_frame, text="Save", command=save_product)
        save_button.grid(row=5, column=0, padx=5, pady=20)

        clear_button = ttk.Button(add_frame, text="Clear", command=lambda: [var.set("") for var in variables])
        clear_button.grid(row=5, column=1, padx=5, pady=20)
    
    # Function to edit a product with the new format
    def edit_product():
        selected_items = treeview.selection()
        if not selected_items:
            messagebox.showwarning("Edit Product", "Please select a product to edit!")
            return

        # Create a new window for editing
        edit_window = Toplevel(root)
        edit_window.title("Edit Product")
        edit_window.geometry("600x400")  # Adjusted size for additional fields
        edit_window.resizable(False, False)
        edit_window.grab_set()  # Make the edit window modal

        # Create a frame inside the edit window
        edit_frame = ttk.Frame(edit_window, padding=10)
        edit_frame.pack(fill=BOTH, expand=True)

        # Function to save the edited product
        def save_edits():
            for idx, item in enumerate(selected_items):
                product_name_var = item_vars[item]['product_name']
                unit_price_var = item_vars[item]['unit_price']
                purchase_cost_var = item_vars[item]['purchase_cost']
                product_print_name_var = item_vars[item]['product_print_name']
                product_print_desc_var = item_vars[item]['product_print_desc']

                # Get the new values from the entry fields
                new_product_name = product_name_var.get().strip()
                new_unit_price = unit_price_var.get().strip()
                new_purchase_cost = purchase_cost_var.get().strip()
                new_product_print_name = product_print_name_var.get().strip()
                new_product_print_desc = product_print_desc_var.get().strip()

                if not new_product_name or not new_unit_price or not new_purchase_cost or not new_product_print_name or not new_product_print_desc:
                    messagebox.showerror("Invalid Input", "All fields are required.")
                    return

                # Update the Treeview
                treeview.item(item, values=(new_product_name, new_unit_price, new_purchase_cost, new_product_print_name, new_product_print_desc))

            # Update the Filtered_data.txt file
            try:
                with open('Filtered_data.txt', 'w') as file:
                    for child in treeview.get_children():
                        values = treeview.item(child, 'values')
                        file.write(f"{values[0]}>{values[1]}>{values[2]}>{values[3]}>{values[4]}\n")
            except Exception as e:
                messagebox.showerror("File Error", f"An error occurred while updating the file: {e}")
                return

            update_item_count()
            messagebox.showinfo("Edit Product", "Product(s) updated successfully!")
            edit_window.destroy()

        # Dictionary to hold StringVars for each selected item
        item_vars = {}

        for idx, item in enumerate(selected_items):
            values = treeview.item(item, 'values')
            product_name, unit_price, purchase_cost, product_print_name, product_print_desc = values

            # Create labels and entry widgets for each selected item
            ttk.Label(edit_frame, text=f"Product {idx + 1}").grid(row=idx * 5, column=0, columnspan=2, pady=(10, 0), sticky='w')

            # Product Name
            ttk.Label(edit_frame, text="Product Name:").grid(row=idx * 5 + 1, column=0, sticky='e', padx=5, pady=5)
            product_name_var = StringVar(value=product_name)
            entry_product_name = ttk.Entry(edit_frame, textvariable=product_name_var)
            entry_product_name.grid(row=idx * 5 + 1, column=1, sticky='w', padx=5, pady=5)

            # Unit Price
            ttk.Label(edit_frame, text="Unit Price:").grid(row=idx * 5 + 2, column=0, sticky='e', padx=5, pady=5)
            unit_price_var = StringVar(value=unit_price)
            entry_unit_price = ttk.Entry(edit_frame, textvariable=unit_price_var)
            entry_unit_price.grid(row=idx * 5 + 2, column=1, sticky='w', padx=5, pady=5)

            # Purchase Cost
            ttk.Label(edit_frame, text="Purchase Cost:").grid(row=idx * 5 + 3, column=0, sticky='e', padx=5, pady=5)
            purchase_cost_var = StringVar(value=purchase_cost)
            entry_purchase_cost = ttk.Entry(edit_frame, textvariable=purchase_cost_var)
            entry_purchase_cost.grid(row=idx * 5 + 3, column=1, sticky='w', padx=5, pady=5)

            # Product Print Name
            ttk.Label(edit_frame, text="Product Print Name:").grid(row=idx * 5 + 4, column=0, sticky='e', padx=5, pady=5)
            product_print_name_var = StringVar(value=product_print_name)
            entry_product_print_name = ttk.Entry(edit_frame, textvariable=product_print_name_var)
            entry_product_print_name.grid(row=idx * 5 + 4, column=1, sticky='w', padx=5, pady=5)

            # Product Print Description
            ttk.Label(edit_frame, text="Product Print Description:").grid(row=idx * 5 + 5, column=0, sticky='e', padx=5, pady=5)
            product_print_desc_var = StringVar(value=product_print_desc)
            entry_product_print_desc = ttk.Entry(edit_frame, textvariable=product_print_desc_var)
            entry_product_print_desc.grid(row=idx * 5 + 5, column=1, sticky='w', padx=5, pady=5)

            # Store the variables in the dictionary
            item_vars[item] = {
                'product_name': product_name_var,
                'unit_price': unit_price_var,
                'purchase_cost': purchase_cost_var,
                'product_print_name': product_print_name_var,
                'product_print_desc': product_print_desc_var
            }

        # Add a Save button
        save_button = ttk.Button(edit_frame, text="Save Changes", command=save_edits)
        save_button.grid(row=len(selected_items) * 5 + 1, column=0, columnspan=2, pady=20)

    # Function to delete a product with the new format
    def delete_product():
        selected_items = treeview.selection()
        if not selected_items:
            messagebox.showwarning("Delete Product", "Please select a product to delete!")
            return

        # Confirm deletion
        confirm = messagebox.askyesno("Delete Product", "Are you sure you want to delete the selected product(s)?")
        if not confirm:
            return

        # Delete selected items from the Treeview
        for item in selected_items:
            treeview.delete(item)

        # Update the Filtered_data.txt file
        try:
            with open('Filtered_data.txt', 'w') as file:
                for child in treeview.get_children():
                    values = treeview.item(child, 'values')
                    file.write(f"{values[0]}>{values[1]}>{values[2]}>{values[3]}>{values[4]}\n")
        except Exception as e:
            messagebox.showerror("File Error", f"An error occurred while updating the file: {e}")
            return

        update_item_count()
        messagebox.showinfo("Delete Product", "Product(s) deleted successfully!")
    
    def preview_selected():
        # Get the manually selected items
        selected_items = treeview.selection()

        # Get the highlighted items (items with the 'highlight' tag)
        highlighted_items = [item for item in treeview.get_children() if 'highlight' in treeview.item(item, 'tags')]

        # Combine both selected and highlighted items (avoiding duplicates)
        all_items = set(selected_items).union(highlighted_items)

        if not all_items:
            messagebox.showwarning("Preview", "Please select or highlight at least one product to preview!")
            return

        # Create a new window to display the selected and highlighted products
        preview_window = Toplevel(root)
        preview_window.title("Preview Selected and Highlighted Products")
        preview_window.geometry("600x400")
        preview_window.resizable(True, False)

        preview_frame = ttk.Frame(preview_window, padding=10)
        preview_frame.pack(fill=BOTH, expand=True)

        # Create a scrollable Text widget to display the selected rows
        preview_text = Text(preview_frame, wrap=NONE)
        preview_text.pack(side=LEFT, fill=BOTH, expand=True)

        scrollbar = Scrollbar(preview_frame, command=preview_text.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        preview_text.config(yscrollcommand=scrollbar.set)

        # Populate the text area with the selected and highlighted rows
        for item in all_items:
            values = treeview.item(item, 'values')
            preview_text.insert(END, f"Product: {values}\n\n")  # Add spacing for better readability

        preview_text.config(state=DISABLED)  # Make the Text read-only
    # Create the main window
    root = Tk()
    window_width = 1400
    window_height = 800


    # Get the screen width and height
    screen_width = root.winfo_screenwidth()
    screen_height =root.winfo_screenheight()

    # Calculate the position to center the window
    x_position = (screen_width - window_width) // 2
    y_position = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")  # Set window dimensions
    root.resizable(True, True) # change back to FALSE when everything is done 

    # Make the root window's grid expandable
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    # Create a frame and place it in the center of the root window
    frm = ttk.Frame(root, padding=50)  # Reduced padding
    frm.grid(column=0, row=0, sticky="nsew")

    # Configure the grid within the frame
    frm.columnconfigure(0, weight=1)
    frm.rowconfigure(1, weight=1)  # Make the row containing the treeview expandable

    # Buttons for Add, Edit, and Delete Product from the text file
    button_frame = Frame(frm)
    button_frame.grid(row=1, column=0, sticky="n", pady=(0, 20))

    # Use a smaller font for buttons
    button_font = ("Arial", 10)
    ttk.Button(button_frame, text="Add", command=add_product, style="Small.TButton").grid(row=0, column=0, padx=2)
    ttk.Button(button_frame, text="Edit", command=edit_product, style="Small.TButton").grid(row=0, column=1, padx=2)
    ttk.Button(button_frame, text="Delete", command=delete_product, style="Small.TButton").grid(row=0, column=2, padx=2)
    ttk.Button(button_frame, text="Preview", command=preview_selected, style="Small.TButton").grid(row=0, column=3, padx=2)

    style = ttk.Style()
    style.configure("Treeview.Heading", font=("Arial",9,"bold"))
    # Create a Scrollbar
    scrollbar = Scrollbar(frm)
    scrollbar.grid(row=1, column=1, sticky="ns",pady=(40,0))

    # Create a Treeview widget with vertical scrollbar
    treeview = ttk.Treeview(frm, yscrollcommand=scrollbar.set, columns=("Product Name", "Unit Price", "Purchase Cost", "Product Print Name", "Product Print Descriptions"), show='headings')
    treeview.grid(row=1, column=0, sticky="nsew",pady=(40,0))

    # Configure scrollbar
    scrollbar.config(command=treeview.yview)



    # Define column headings
    #treeview.heading("Select", text="Select", anchor='w')
    treeview.heading("Product Name", text="Product Name")
    treeview.heading("Unit Price", text="Unit Price")
    treeview.heading("Purchase Cost", text="Purchase Cost")
    treeview.heading("Product Print Name", text="Product Print Name")
    treeview.heading("Product Print Descriptions", text="Product Print Descriptions")

    # Set the column widths
    # Set the column widths
    #treeview.column("Select", width=100, anchor='center', stretch=False)
    treeview.column("Product Name", width=200, anchor='w')
    treeview.column("Unit Price", width=100, anchor='center')
    treeview.column("Purchase Cost", width=100, anchor='center')
    treeview.column("Product Print Name", width=200, anchor='w')
    treeview.column("Product Print Descriptions", width=400, anchor='w')

    # Read items from the Filtered_data.txt file and insert them into the Treeview
    item_count = 0
    with open('Filtered_data.txt', 'r') as file:
        for line in file:
            item_count += 1  # Count each line inserted
            # Split line by '>' to separate product details
            parts = line.strip().split('>')
            if len(parts) == 5:
                product_name, unit_price, purchase_cost, product_print_name, product_print_descriptions = parts
            else:
                continue  # Skip lines that don't have exactly 5 fields

            # Insert the product details into the Treeview
            treeview.insert("", "end", values=(product_name, unit_price, purchase_cost, product_print_name, product_print_descriptions))


        # Create a label to display the number of items with specified font
        labelItem = Label(frm, text=f"Number of items: {item_count}", font=("Arial", 12))  # Reduced font size
        labelItem.grid(column=0, row=3, sticky="ew", padx=5, pady=2)


        # Add a quit button
        ttk.Button(frm, text="Quit", command=root.destroy, style="Small.TButton").grid(column=0, row=4, sticky="ew", padx=5, pady=2)
        ttk.Label(frm,text='NOTE :  To select MULTIPLE ITEMS , please hold the CTRL key and select the item you want',font='Helvetica-BOLD').grid(column=0, row= 5 , sticky="s",padx=5,pady=5)


        # Search Bar 
        # Create a style for highlighting the rows
        style = ttk.Style()
        style.configure("Treeview.Highlight", background="yellow")  # Yellow highlight for matching rows

        # Configure the Treeview to use the highlight style for matching rows
        treeview.tag_configure('highlight', background="yellow")

        # Add a frame for the search bar above the Treeview
        search_frame = Frame(frm)
        search_frame.grid(row=0, column=0, sticky="n", pady=(0, 30)) # pady = x , y

        # Search input
        search_var = StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var, width=80)
        search_entry.grid(row=0, column=0, padx=5, pady=5)

        # Placeholder text 
        search_entry.insert(0, "Search by Product Name, Unit Price, Purchase Cost, Print Name, or Descriptions")

        # Search and Clear buttons
        ttk.Button(search_frame, text="Search", command=search_and_highlight).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(search_frame, text="Clear", command=clear_search_and_highlight).grid(row=0, column=2, padx=5, pady=5)

        # Bind Enter key to the search function
        search_entry.bind('<Return>', lambda event: search_and_highlight())

        # Create a style for smaller buttons
        style = ttk.Style()
        style.configure("Small.TButton", padding=2, font=button_font)


        # Row selection highlighting
        #treeview.bind("<ButtonRelease-1>", toggle_select)
        root.mainloop()

## MAIN WINDOWS ##
# Main windows for editing / adding products into the VTiger CRM website
def main_windows(product_info, driver=NONE):
    # This function is to add a new row (new product) under the existing row  
    def add_products():
        #product_element = WebDriverWait(driver, 10).until(
        #    EC.element_to_be_clickable((By.XPATH, "//span[contains(@class, 'text-primary') and contains(@class, 'c-pointer') and contains(text(), 'Products')]"))
        #)
        # Need to click twice in order to open a new row
        #product_element.click()
        #product_element.click()

        # Switch back to the default content after interaction
        driver.switch_to.default_content()
        # Windows for adding new products in the last row 
        def edit_comment_name():
            def search_comment_item(event):
                search_query = search_comment_entry.get().lower()
                comment_listbox.delete(0, END)

                if search_query:
                    with open("Filtered_data.txt", "r") as file:
                        found = False
                        for line in file:
                            item_name = line.split(">")[0]  # Get only the first column
                            if search_query in item_name.lower():
                                comment_listbox.insert(END, item_name)
                                found = True
                        if not found:
                            comment_listbox.insert(END, "No results found")
                else:
                    comment_listbox.insert(END, "No results found")

            def select_comment_items():
                selected_items = comment_listbox.curselection()
                for item_index in selected_items:
                    item_name = comment_listbox.get(item_index)
                    if item_name != "No results found":
                        text_comment.insert(END, item_name + "\n")  # Add item to text_comment
                root_comment.destroy()  # Close the edit window

            def cancel_comment_edit():
                root_comment.destroy()

            # Create the edit comment window
            root_comment = Toplevel(root_new_row)
            root_comment.title("Add / Delete Comment Items")
            window_width_comment = 550
            window_height_comment = 400

            screen_width_comment = root_comment.winfo_screenwidth()
            screen_height_comment = root_comment.winfo_screenheight() 

            x_position = (screen_width_comment - window_width_comment) // 2
            y_position = (screen_height_comment - window_height_comment) // 2
            root_comment.geometry(f"{window_width_comment}x{window_height_comment}+{x_position}+{y_position}")
            root_comment.resizable(False, False)
            root_comment.transient(root_new_row)
            root_comment.grab_set()

            # Search bar
            Label(root_comment, text="Search Product items", font="BOLD").pack(pady=10)
            search_comment_entry = Entry(root_comment, width=60)
            search_comment_entry.pack(pady=5)
            search_comment_entry.bind("<KeyRelease>", search_comment_item)

            # Listbox for displaying search results
            comment_listbox = Listbox(root_comment, width=60, height=10, selectmode=MULTIPLE)
            comment_listbox.pack(pady=5)

            # Frame for Select and Cancel buttons
            button_frame = Frame(root_comment)
            button_frame.pack(pady=10)

            # Select and Cancel buttons
            select_button = Button(button_frame, text="Select", command=select_comment_items)
            select_button.grid(row=0, column=0, padx=5)

            cancel_button = Button(button_frame, text="Cancel", command=cancel_comment_edit)
            cancel_button.grid(row=0, column=1, padx=5)


        def search_item(event):
            search_query = search_entry.get().lower()
            suggestions_listbox.delete(0, END)

            if search_query:  # Proceed only if there is a search query
                with open("Filtered_data.txt", "r") as file:
                    for line in file:
                        item_name = line.split(">")[0]  # Get only the first column
                        if search_query in item_name.lower():
                            suggestions_listbox.insert(END, item_name)  # Add to Listbox if query matches
            else:
                suggestions_listbox.insert(END, "No results found")  # Show if no matching items

        # Function to handle item selection from the Listbox
        def select_item(event):
            try:
                selected_item = suggestions_listbox.get(suggestions_listbox.curselection())
                search_entry.delete(0, END)
                search_entry.insert(0, selected_item)  # Display the selected item in the search bar
            except TclError:
                pass  # Ignore if there is no selection

        # Function to open the item edit window
        def edit_item_name():
            def cancel_edit():
                root_edit.destroy()  # Close the edit window

            def select_and_close():
                try:
                    selected_item = suggestions_listbox.get(suggestions_listbox.curselection())
                    entry_item_name.config(text=selected_item)  # Set the selected item to the main label
                    root_edit.destroy()  # Close the edit window
                except TclError:
                    messagebox.showerror("Selection Error", "Please select a valid item from the list.")

            root_edit = Toplevel(root_new_row)
            root_edit.title("Select new item's name")
            window_width_edit = 550
            window_height_edit = 400

            screen_width_edit = root_edit.winfo_screenwidth()
            screen_height_edit = root_edit.winfo_screenheight() 

            x_position = (screen_width_edit - window_width_edit) // 2
            y_position = (screen_height_edit - window_height_edit) // 2
            root_edit.geometry(f"{window_width_edit}x{window_height_edit}+{x_position}+{y_position}")
            root_edit.resizable(False, False)

            root_edit.transient(root_new_row)  # Keep the edit window on top of the main window
            root_edit.grab_set()       # Disable interaction with the main window while edit window is open

            # Search bar
            Label(root_edit, text="Search Product's name", font="BOLD").pack(pady=10)
            global search_entry  # Make search_entry global to access in search_item function
            search_entry = Entry(root_edit, width=60)
            search_entry.pack(pady=5)
            search_entry.bind("<KeyRelease>", search_item)  # Bind key release event for dynamic search

            # Listbox for suggestions
            global suggestions_listbox
            suggestions_listbox = Listbox(root_edit, width=60, height=10)
            suggestions_listbox.pack(pady=5)
            suggestions_listbox.bind("<<ListboxSelect>>", select_item)  # Bind selection event

            # Frame for "Select" and "Cancel" buttons
            button_frame = Frame(root_edit)
            button_frame.pack(pady=10)

            # "Select" button
            select_button = Button(button_frame, text="Select", command=select_and_close)
            select_button.grid(row=0, column=0, padx=5)

            # "Cancel" button
            cancel_button = Button(button_frame, text="Cancel", command=cancel_edit)
            cancel_button.grid(row=0, column=1, padx=5)

        # This function is to clear the comment's text 
        def clear_comment():
            text_comment.delete("1.0", END)  # Clear all text in the text_comment widget

        # This function is to get the last row ( new product details) after the confirm button is pressed and send to CRM
        def send_to_crm_after_confirm(root_new_row):
            new_quantity = spinbox_quantity.get()
            new_item_name = entry_item_name.cget("text")
            new_comment = text_comment.get("1.0",END).strip()
            new_unit_selling_price = entry_unit_selling_price.get()

            # Validate if Quantity and Item Name are not empty
            if not new_quantity or new_quantity == '0':  # Ensure quantity is not empty or zero
                messagebox.showerror("Input Error", "Quantity cannot be empty or zero!")
                return  # Stop further execution if validation fails

            if not new_item_name.strip():  # If Item Name is just whitespace
                messagebox.showerror("Input Error", "Item Name cannot be empty!")
                return  # Stop further execution if validation fails
            
            # For debugging purpose 
            print(f"Quantity: {new_quantity}")
            print(f"Item Name: {new_item_name}")
            print(f"Comment: {new_comment}")
            print(f"Unit Selling Price: {new_unit_selling_price}")

            # Switch back to default content and locate the "Products" element
            driver.switch_to.default_content()
            product_element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//span[contains(@class, 'text-primary') and contains(@class, 'c-pointer') and contains(text(), 'Products')]"))
            )
        
            product_element.click() 
            try:
                # Check if a new row is added by waiting for it to appear
                WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, "//tr[contains(@id, 'lineItemRow')]"))
                )
            except:
                # If a new row didn't appear within 3 seconds, click again
                product_element.click()

            wait = WebDriverWait(driver, 10)
            rows = wait.until(EC.presence_of_all_elements_located(
                (By.XPATH, "//tr[contains(@id, 'lineItemRow')]")
            ))
            row_number = len(rows)  # This will be the last row number
            base_id = f'lineItemRow{row_number}'

            # Handle item name insertion following the specified steps
            # Step 1 & 3: Click on the dropdown field to activate the search input
            dropdown_field = wait.until(EC.element_to_be_clickable((
                By.XPATH, 
                f"//tr[@id='{base_id}']//span[contains(@class, 'select2-selection__rendered')]"
            )))
            dropdown_field.click()

            # Step 2 & 4: Type the item name into the search input field
            search_input = wait.until(EC.presence_of_element_located((
                By.XPATH, 
                f"//tr[@id='{base_id}']//input[@class='select2-search__field']"
            )))
            search_input.send_keys(new_item_name)

            # Click the first result in the dropdown menu
            first_result = wait.until(EC.element_to_be_clickable((
                By.XPATH, "//li[contains(@class, 'select2-results__option--highlighted')]"
            )))
            first_result.click()

            # Fill in comment
            comment_element = wait.until(EC.presence_of_element_located((
                By.XPATH, f"//textarea[@name='comment{row_number}']"
            )))
            comment_element.clear()
            comment_element.send_keys(new_comment)

            # Fill in unit selling price
            price_element = wait.until(EC.presence_of_element_located((
                By.XPATH, 
                f"//tr[@id='{base_id}']//input[@test-attr='listPrice{row_number}']"
            )))
            price_element.clear()
            price_element.send_keys(new_unit_selling_price)
            # Also set the test-attr-price attribute
            driver.execute_script("arguments[0].setAttribute('test-attr-price', arguments[1])", 
                                price_element, new_unit_selling_price)
            
             # Fill in quantity last to prevent it from being reset
            time.sleep(1)  # Add a small delay to ensure other fields are settled
            quantity_element = wait.until(EC.presence_of_element_located((
                By.XPATH, 
                f"//tr[@id='{base_id}']//input[@name='qty{row_number}']"
            )))
            quantity_element.clear()
            quantity_element.send_keys(new_quantity)
    
            # Add an extra step to ensure the quantity stays set
            driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'))", 
                                quantity_element, new_quantity)
            
            fetched_edited_last_row = extract_products_info(driver)
            if update_treeview:
                update_treeview(fetched_edited_last_row)

            messagebox.showinfo("Item added to CRM","The item has been added into CRM successfully.")
            root_new_row.destroy()

        root_new_row = Tk()
        root_new_row.title("Enter new product's details")
        window_width = 640
        window_height = 400

        # Get the screen width and height
        screen_width = root_new_row.winfo_screenwidth()
        screen_height = root_new_row.winfo_screenheight() 

        # Calculate the position to center the window
        x_position = (screen_width - window_width) // 2
        y_position = (screen_height - window_height) // 2
        root_new_row.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
        root_new_row.resizable(False, False)

        # Set the main frame for the form
        edit_main_frame = Frame(root_new_row)
        edit_main_frame.grid(padx=10, pady=10)

        # Quantity (using Spinbox for increment and decrement)
        label_quantity = Label(edit_main_frame, text="Quantity")
        label_quantity.grid(row=0, column=0, sticky=E, padx=5, pady=25)
        spinbox_quantity = Spinbox(edit_main_frame, from_=1, to=99999, width=10)
        spinbox_quantity.grid(row=0, column=1, sticky=W, padx=5, pady=25)

        # Item Name
        label_item_name = Label(edit_main_frame, text="Item Name")
        label_item_name.grid(row=1, column=0, sticky=E, padx=5, pady=20)
        entry_item_name = Label(edit_main_frame, text="", font="BOLD",wraplength=250)
        entry_item_name.grid(row=1, column=1, sticky=W, padx=5, pady=20)

        # Add edit button next to Item Name
        button_item_name = Button(edit_main_frame, text="EDIT", width=5, font="BOLD", command=edit_item_name,justify=LEFT)
        button_item_name.grid(row=1, column=3, sticky=W, padx=5)

        # Comment
        label_comment = Label(edit_main_frame, text="Comment")
        label_comment.grid(row=2, column=0, sticky=E, padx=5, pady=2)

        # Use Text widget for multiline input
        text_comment = Text(edit_main_frame, width=30, height=4, wrap="word")
        text_comment.grid(row=2, column=1, sticky=W, padx=5, pady=2)

        # Set up a scroll bar for the Text widget
        scrollbar_comment = Scrollbar(edit_main_frame, command=text_comment.yview)
        scrollbar_comment.grid(row=2, column=2, sticky="ns")
        text_comment.config(yscrollcommand=scrollbar_comment.set)

        # Add folder button next to Comment scrollbar
        button_comment = Button(edit_main_frame, text="+", width=2, font="BOLD",command=edit_comment_name)
        button_comment.grid(row=2, column=3, sticky=E, padx=10)

        trash_comment = Button(edit_main_frame, text="🗑", width=2, font="BOLD", command=clear_comment)
        trash_comment.grid(row=2, column=4, sticky=W, padx=0)

        # Unit Selling Price
        label_unit_selling_price = Label(edit_main_frame, text="Unit Selling Price")
        label_unit_selling_price.grid(row=3, column=0, sticky=E, padx=5, pady=2)
        entry_unit_selling_price = Entry(edit_main_frame, width=30)
        entry_unit_selling_price.grid(row=3, column=1, sticky=W, padx=5, pady=2)

        # Add a frame for the Cancel and Confirm buttons at the bottom
        button_frame = Frame(root_new_row)
        button_frame.grid(row=4, column=0, sticky=E, padx=0, pady=10)

        # Cancel and Confirm buttons
        button_cancel = Button(button_frame, text="Cancel", width=10, command=root_new_row.destroy)
        button_cancel.grid(row=0, column=0, padx=5)

        button_confirm = Button(button_frame, text="Confirm", width=10, command = lambda : send_to_crm_after_confirm(root_new_row))
        button_confirm.grid(row=0, column=1, padx=5)

        edit_main_frame.grid_columnconfigure(1,weight =1)
        # Run the main event loop
        root_new_row.mainloop()

    def update_treeview(data):
        """Clears and refreshes the Treeview with the latest data from product_info."""
        product_info[:] = data  # Update product_info in-place
        for item in table.get_children():
            table.delete(item)  # Clear all rows from the Treeview

        for i, product in enumerate(product_info):
            item_name = product['item_name']
            if 'subproducts' in product and product['subproducts']:
                subproducts_str = "\n".join(product['subproducts'])
                item_name += f"\n{subproducts_str}"
            if 'comment' in product and product['comment']:
                item_name += f"\n{product['comment']}"

            table.insert("", "end", values=(
                i + 1, product['quantity'], item_name,
                product['unit_selling_price'], product['total_purchase_cost'],
                product['unit_purchase_cost'], product['total_after_discount'],
                product['margin'], product['tax'], product['net_price']
            ))
    
    # This Function is to prevent the Treeview's column from resizing 
    # TODO : TO RESIZE the Treeview's , just comment this function so that it will resize again 
    def handle_click(event):
        # Disable resizing by blocking the separator clicks
        if table.identify_region(event.x, event.y) == "separator":
            return "break"

    def on_edit_button_click():
        # Get the selected item
        selected_item = table.selection()
        if selected_item:
            # Get the values of the selected row 
            row_data = table.item(selected_item[0], 'values')
            
            # Get the index (No.) from the selected row
            selected_index = int(row_data[0]) - 1  # Subtract 1 since display index starts from 1
            
            # Get the corresponding product from product_info
            selected_product = product_info[selected_index]
            
            # Create new row data using the original product_info data
            new_row_data = [
                row_data[0],  # Index 0 - No (keep original display number)
                selected_product['quantity'],  # Index 1 - Quantity
                selected_product['item_name'],  # Index 2 - Item Name
                selected_product.get('subproducts', []),  # Index 3 - Subproducts (empty list if doesn't exist)
                selected_product.get('comment', ''),  # Index 4 - Comment (empty string if doesn't exist)
                selected_product['unit_selling_price'],  # Index 5 - Unit Selling Price
                selected_product['total_purchase_cost'],  # Index 6 - Total Purchase Cost
                selected_product['unit_purchase_cost'],  # Index 7 - Unit Purchase Cost
                selected_product['total_after_discount'],  # Index 8 - Total
                selected_product['margin'],  # Index 9 - Margin
                selected_product['tax'],  # Index 10 - Tax
                selected_product['net_price']  # Index 11 - Net Price
            ]
            
            # Format subproducts for display
            subproducts_display = ""
            if new_row_data[3]:  # If subproducts exist
                subproducts_display = "\n".join(f"{subproduct}" for subproduct in new_row_data[3])
            
            # For debugging purpose 
            print("Selected Product Data:")
            print(f"Index 0 (No): {new_row_data[0]}")
            print(f"Index 1 (Quantity): {new_row_data[1]}")
            print(f"Index 2 (Item Name): {new_row_data[2]}")
            print(f"Index 3 (Subproducts):{subproducts_display}" if subproducts_display else "Index 3 (Subproducts): None")
            print(f"Index 4 (Comment): {new_row_data[4]}")
            print(f"Index 5 (Unit Selling Price): {new_row_data[5]}")
            print(f"Index 6 (Total Purchase Cost): {new_row_data[6]}")
            print(f"Index 7 (Unit Purchase Cost): {new_row_data[7]}")
            print(f"Index 8 (Total): {new_row_data[8]}")
            print(f"Index 9 (Margin): {new_row_data[9]}")
            print(f"Index 10 (Tax): {new_row_data[10]}")
            print(f"Index 11 (Net Price): {new_row_data[11]}")
            
            # Now you can pass new_row_data to edit_prod_crm or handle it as needed
            edit_prod_crm(new_row_data,product_info,driver,update_treeview)

    def on_save_and_close_click():
        # Run save_and_close in a separate thread to avoid blocking Tkinter UI
        threading.Thread(target=save_and_close, args=(driver,)).start()

    root = Tk()
    root.title("Extracted data fom VTiger")
    window_width = 1400   # Set Window width HERE
    window_height = 800   # Set Window height HERE

    # Get the screen width and height
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight() 

    # Calculate the position to center the window
    x_position = (screen_width - window_width) // 2
    y_position = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")  # Set window dimensions
    root.resizable(True, True) # change back to FALSE when everything is done 

    # Make the root window's grid expandable 
    root.columnconfigure(0, weight =1)
    root.rowconfigure(0, weight =1)

    # Set padding value here
    # Create a frame and place it in the center of the root window
    frm = ttk.Frame(root, padding=50)
    frm.grid(column=0, row=0, sticky="nsew")

    # Set custom style for the Treeview header to increase its height
    style = ttk.Style()
    style.configure("Treeview.Heading", font=("Arial", 10, "bold"), padding=[10, 1])  # Increase font and add padding
    style.configure("Treeview", rowheight =150) # Set font and row height for data rows 

    # Add a table with columns
    columns = ("No", "Quantity", "Item Name", "Unit Selling Price", "Item Total Purchase Cost", "Item Unit Purchase Cost","Total","Margin","Tax","Net Price")
    table = ttk.Treeview(frm, columns=columns, show="headings", style="Treeview")

    # Define column headers and set fixed column widths
    column_widths = [50, 100, 350, 180, 250, 220, 100, 100, 100, 120]  # Example widths for each column

    for col, width in zip(columns, column_widths):
        table.heading(col, text=col)
        table.column(col, anchor="center", width=width, minwidth=width, stretch=False)  # Set fixed width

    # Insert rows into the Treeview
    #for i, product in enumerate(product_info):
    #    # Check if subproducts or comment exist and format them accordingly
    #    item_name = product['item_name']
    #    if 'subproducts' in product and product['subproducts']:
    #        # Join subproducts list into a string, with each item on a new line
    #        subproducts_str = "\n".join(product['subproducts'])  # This joins each item with a newline
    #        item_name += f"\n{subproducts_str}"
    #    if 'comment' in product and product['comment']:
    #       item_name += f"\n{product['comment']}"
        
        # Insert row with formatted item_name
    #    table.insert("", "end", values=(i + 1, product['quantity'], item_name,
    #                                    product['unit_selling_price'], product['total_purchase_cost'],
    #                                    product['unit_purchase_cost'], product['total_after_discount'],
    #                                    product['margin'], product['tax'], product['net_price']))
    # Add table to the grid
    table.grid(column=0, row=0, sticky="nsew")
    update_treeview(product_info)
    # Create a horizontal scrollbar
    h_scrollbar = ttk.Scrollbar(frm, orient="horizontal", command=table.xview)
    h_scrollbar.grid(row=1, column=0, sticky="ew")  # Place it below the Treeview
    table.configure(xscrollcommand=h_scrollbar.set)  # Link the scrollbar to the Treeview

    # Disable resizing by binding click events to prevent resizing by dragging
    # To resize it , comment the line under this comment
    table.bind("<Button-1>", handle_click)

    # Allow frame and table to expand
    frm.columnconfigure(0, weight=1)
    frm.rowconfigure(0, weight=1)

    button_frame = ttk.Frame(frm)
    button_frame.grid(column = 0, row=2, pady =10 ,sticky="ew")
    #Add "Add", "Edit", and "Close" buttons to the button frame under the table
    add_button = ttk.Button(button_frame, text="Add Product(s)",command=add_products)
    add_button.grid(column=0, row=0, padx=10)
    edit_button = ttk.Button(button_frame, text="Edit Product",command=on_edit_button_click)
    edit_button.grid(column=1, row=0, padx=10)
    manage_button = ttk.Button(button_frame, text="Manage Items", command=select_prod_windows) # This button is to manage the items / products in the textfile
    manage_button.grid(column=2, row=0, padx=20)
    close_button = ttk.Button(button_frame, text="Save & Close",command = on_save_and_close_click)
    close_button.grid(column=3, row=0, padx=10)
    
    button_frame.columnconfigure(0, weight=1)
    button_frame.columnconfigure(1, weight=1)
    button_frame.columnconfigure(2, weight=1)
    button_frame.columnconfigure(3, weight=1)

    # Run the main event loop
    root.mainloop()

# Function to send the keys (username, password and quoteID) to the Selenium
def on_submit(root):
    username = username_entry.get() # Get the username field's value
    password = password_entry.get() # Get the password field's value
    quote_id = quote_id_entry.get().upper() # Get the quote ID field's value

    if not username or not password or not quote_id:
        messagebox.showerror("Input Error", "Please enter all fields.")
        return

    # Run fetch_line_item_rows in a separate thread to avoid freezing the GUI
    threading.Thread(target=fetch_line_item_rows, args=(username, password, quote_id)).start()
    root.destroy()
        
# This is the GUI windows for login into the VTiger CRM system in the beginning 
def login_gui():
    root = tk.Tk()
    root.title("VTiger Login")
    root.iconbitmap("crm.ico")
    window_width = 400   # Set Window width HERE
    window_height = 300   # Set Window height HERE

    # Get the screen width and height
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight() 

    # Calculate the position to center the window
    x_position = (screen_width - window_width) // 2
    y_position = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")  # Set window dimensions
    root.resizable(True, True) # change back to FALSE when everything is done 

    # Username
    tk.Label(root, text="Username:").pack(pady=5)
    global username_entry
    username_entry = tk.Entry(root, width=30)
    username_entry.pack(pady=5)

    # Password
    tk.Label(root, text="Password:").pack(pady=5)
    global password_entry
    password_entry = tk.Entry(root, width=30)
    password_entry.pack(pady=5)

    # Quote ID
    tk.Label(root, text="Quote ID:").pack(pady=5)
    global quote_id_entry
    quote_id_entry = tk.Entry(root, width=30)
    quote_id_entry.pack(pady=5)

    # Submit Button
    submit_button = tk.Button(root, text="Submit", command=on_submit)
    submit_button.pack(pady=20)

    # Bind Enter key to submit action
    root.bind("<Return>", lambda event: on_submit(root))

    root.mainloop()

# This function is to get the products' information from each row ( You can even get each row using iteration for fetching data in the table for each row)
def get_product_info(driver, row_number):
    """
    Extract product information for a specific row number with explicit waits
    """
    wait = WebDriverWait(driver, 20)
    base_id = f"lineItemRow{row_number}"
    product_info = {}
    
    try:
        # Wait for the row to be present
        wait.until(EC.presence_of_element_located((By.ID, base_id)))
        

        # Quantity
        quantity_element = wait.until(EC.presence_of_element_located((
            By.XPATH, 
            f"//tr[@id='{base_id}']//input[@name='qty{row_number}']"
        )))
        product_info['quantity'] = quantity_element.get_attribute('value')
        
        # Item Name
        item_name_element = wait.until(EC.presence_of_element_located((
            By.XPATH, 
            f"//tr[@id='{base_id}']//input[@readonly='readonly' and contains(@class, 'form-control bg-white')]"
        )))
        product_info['item_name'] = item_name_element.get_attribute('title')
        
        # Extract the subProducts line 
        subproducts_element = driver.find_element(By.CSS_SELECTOR, f"#lineItemRow{row_number} .subProducts")
        if subproducts_element.is_displayed():
            subproducts = []
            for subproduct_div in subproducts_element.find_elements(By.CSS_SELECTOR, 'div'):
                subproduct_text = subproduct_div.text.strip()
                if subproduct_text:
                    subproducts.append(subproduct_text)
            product_info['subproducts'] = subproducts

        # Extract the comment's line 
        comment_element = wait.until(EC.presence_of_element_located((
            By.XPATH, f"//textarea[@name='comment{row_number}']"
        )))
        product_info['comment'] = comment_element.get_attribute('value')
        
        # Unit Selling Price
        price_element = wait.until(EC.presence_of_element_located((
            By.XPATH, 
            f"//tr[@id='{base_id}']//input[@test-attr='listPrice{row_number}']"
        )))
        product_info['unit_selling_price'] = price_element.get_attribute('test-attr-price')
        
        # Item Total Purchase Cost
        total_cost_element = wait.until(EC.presence_of_element_located((
            By.XPATH, 
            f"//tr[@id='{base_id}']//input[@type='number' and @name='purchase_cost{row_number}']"
        )))
        product_info['total_purchase_cost'] = total_cost_element.get_attribute('value')
        
        # Item Unit Purchase Cost
        unit_cost_element = wait.until(EC.presence_of_element_located((
            By.XPATH, 
            f"//tr[@id='{base_id}']//input[@name='unit_purchase_cost{row_number}']"
        )))
        product_info['unit_purchase_cost'] = unit_cost_element.get_attribute('value')
        
        # Total After Discount
        total_discount_element = wait.until(EC.presence_of_element_located((
            By.ID, f"totalAfterDiscount{row_number}"
        )))
        product_info['total_after_discount'] = total_discount_element.text
        
        # Margin
        margin_element = wait.until(EC.presence_of_element_located((
            By.ID, f"margin{row_number}"
        )))
        product_info['margin'] = margin_element.text
        
        # Tax
        tax_element = wait.until(EC.presence_of_element_located((
            By.XPATH, f"//td[@id='taxTotal{row_number}']//button"
        )))
        product_info['tax'] = tax_element.text
        
        # Net Price
        net_price_element = wait.until(EC.visibility_of_element_located((By.ID, f"netPrice{row_number}")))
        product_info['net_price'] = net_price_element.text
        
        return product_info
        
    except TimeoutException as e:
        print(f"Timeout waiting for element in row {row_number}: {str(e)}")
        return None
    except NoSuchElementException as e:
        print(f"Element not found in row {row_number}: {str(e)}")
        return None
    except Exception as e:
        print(f"Error extracting information for row {row_number}: {str(e)}")
        return None

# This is a helper method for fetching every row 
def extract_products_info(driver):
    try:
        # First ensure we're on the right page and the table is loaded
        wait = WebDriverWait(driver, 20)
        
        # Wait for at least one row to be present
        print("Waiting for line items to load...")
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "lineItemRow")))
        
        # Give a short pause to ensure all elements are properly loaded
        time.sleep(2)
        
        print("Line items found, extracting product information...")

        # Find all rows
        rows = driver.find_elements(By.CLASS_NAME, "lineItemRow")
        products = []

        # Extract information for each product row found
        for row_element in rows:
            row_number = int(row_element.get_attribute('id').replace('lineItemRow', ''))
            print(f"Extracting information for product {row_number}...")            
            # Scroll to the product row
            driver.execute_script("arguments[0].scrollIntoView(true);", row_element)

            
            # Now get the product info
            product_info = get_product_info(driver, row_number)
            if product_info:
                products.append(product_info)
                print(f"Successfully extracted information for product {row_number}")
            else:
                print(f"Failed to extract information for product {row_number}")
        
        return products
    except TimeoutException as e:
        print(f"Timeout waiting for line items: {str(e)}")
        return []
    except Exception as e:
        print(f"Error in extract_products_info: {str(e)}")
        return []

# This function is for clicking neccessary element like some elements that allow the loading of other element due to the CRM is a LAZY loading website
def fetch_line_item_rows(username, password, quote_id):
    driver = None
    try:
        options = uc.ChromeOptions()
        # If you want to go invisible , add '--headless'
        options.add_argument('--start-minimized')
        
        # Initialize the undetected-chromedriver with options
        print("Initializing Chrome driver...")
        driver = uc.Chrome(options=options)
        
        # Load the webpage
        print("Navigating to login page...")
        driver.get('https://crmaccess.vtiger.com/log-in/')
        time.sleep(10)
        
        # Login (selenium will first find the username and password element 1st then , it will send the username and password to the form )
        print("Logging in...")
        email_field = driver.find_element(By.NAME, "username")
        password_field = driver.find_element(By.NAME, "password")
        email_field.send_keys(username)
        password_field.send_keys(password)
        password_field.send_keys(Keys.RETURN)
        time.sleep(25)
        
        # Navigate to Quotes
        print("Navigating to Quotes section...")
        driver.find_element(By.CSS_SELECTOR, "#menu").click()
        search_field = driver.find_element(By.CSS_SELECTOR, 'input.form-control.w-100.h-100.rounded.border-grey-1')
        search_field.send_keys("Quotes")
        driver.find_element(By.CSS_SELECTOR, 'span[title="Quotes"]').click()
        time.sleep(5)
        
        # Search for quote
        print(f"Searching for quote ID: {quote_id}")
        rows = driver.find_elements(By.CSS_SELECTOR, 'div.textOverflowEllipsis span[title]')
        quote_found = False
        
        # First try direct search in visible rows
        for span in rows:
            if span.get_attribute("title") == quote_id:
                link = span.find_element(By.TAG_NAME, 'a')
                link.click()
                quote_found = True
                print("Quote found in visible rows")
                break

        # If not found, try searching
        if not quote_found:
            print("Quote not found in visible rows, trying search...")
            search_input = driver.find_element(By.CSS_SELECTOR, 'input.form-control[placeholder="Search"]')
            search_input.clear()
            search_input.send_keys(quote_id)
            search_input.send_keys(Keys.RETURN)
            time.sleep(10)

            search_results = driver.find_elements(By.CSS_SELECTOR, 'div.textOverflowEllipsis span[title]')
            for span in search_results:
                if span.get_attribute("title") == quote_id:
                    link = span.find_element(By.TAG_NAME, 'a')
                    link.click()
                    quote_found = True
                    print("Quote found through search")
                    break

        if not quote_found:
            print("Quote ID not found.")
            return None

        print("Navigating to quote details...")
        time.sleep(2)
        
        # Click Details tab
        print("Clicking Details tab...")
        details_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "span[title='Details']"))
        )
        details_button.click()
        time.sleep(3)

        # Scroll to and click grand total
        try:
            print("Scrolling to grand total...")
            grand_total_div = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "grandTotal"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", grand_total_div)
            time.sleep(2)
            print("Clicking grand total...")
            driver.execute_script("arguments[0].click();", grand_total_div)
            time.sleep(2)
            

            print("Scrolling to section name input...")
            section_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((
                By.CSS_SELECTOR, 
                'input[type="text"][placeholder="Section Name"].form-control.mr-2'
                ))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", section_input)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", section_input)
            time.sleep(1)

            # Extract product information
            print("Extracting product information...")
            products_info = extract_products_info(driver)
            
            if products_info:
                print("Successfully extracted product information:")
                for i, product in enumerate(products_info, 1):
                    print(f"\nProduct {i} Information:")
                    for key, value in product.items():
                        print(f"{key}: {value}")
                return products_info
                
            else:
                print("No product information was extracted")
                return None
            
            
        except Exception as e:
            print(f"Error accessing product information: {str(e)}")
            return None

    except Exception as e:
        print(f"Error in fetch_line_item_rows: {str(e)}")
        return None
    
    # If you need to add function or anything , add here after the extraction process is done 
    finally:
        if driver:
            try:
                time.sleep(5)
                
                # This is to pass the products_info that was extracted by the extraction function and pass it to the main_windows to display it on the TreeView
                if products_info:
                    main_windows(products_info,driver)

            except Exception as e:
                print(f"Error closing driver: {str(e)}")
                driver.quit()

# Main class
if __name__ == "__main__":
        login_gui()