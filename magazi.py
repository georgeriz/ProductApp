from Tkinter import *
from tkFileDialog import askopenfilename
from codecs import open
from ttk import *
import tkMessageBox
from os import startfile
from myProduct import *
from myWordChange import *
from myTags import TAGS as labels
TAGS = {}
for key in labels:
    TAGS[key] = labels[key][0]

app_icon = "data_files/bolt.ico"
data_location = "data_files/data_location.txt"
data_start = "data_files/data_start.txt"
help_file_location = "data_files\help.txt"

# helper functions
def pool(filename, info):
    my_list = []
    try:
        reader = ExcelReader(filename, info)
        my_list = reader.run()
    except:
        tkMessageBox.showinfo(TAGS["import_error_title"], TAGS["import_error_message"])
    return my_list

def search_by_name(name):
    result_list = _all_product_list
    full_length = len(result_list)
    name_list = name.split()
    for term in name_list:
        term = _word_change.transform(term)
        result_list = _search_term(term, result_list)
    if len(result_list) == full_length:
        return []
    return result_list

def _search_term(term, my_list):
    res = []
    for product in my_list:
        try:
            product_names = product.get_name().split()
            for product_name_term in product_names:
                term = term.lower().strip("(").strip(")")
                product_name_term = product_name_term.lower().strip("(").strip(")")
                if product_name_term.startswith(term):
                    res.append(product)
                    break
        except:
            continue
    if len(res) == 0:
        return my_list
    return res

def search_by_code(code):
    result_list = []
    my_list = _all_product_list
    code = code.strip()
    #Codes are in the form: 123,456,7890
    code = code.replace(".", ",")
    code = code.replace(" ", ",")
    if len(code) > 3 and code[3] != ",":
        code = code[:3] + "," + code[3:]
    if len(code) > 7 and code[7] != ",":
        code = code[:7] + "," + code[7:]
    for product in my_list:
        if product.get_code().startswith(code):
            result_list.append(product)
    return result_list

def _insert_result(result):
    row_frame = Frame(disposable_frame)
    row_frame.pack(fill = X, expand = True)
    label1 = Label(row_frame, text = result.get_code(), width = 12, anchor = W)
    label1.pack(side = LEFT)
    label1.bind("<Button-1>", DescriptionWrap(result).fetch_description)
    label2 = Label(row_frame, text = result.get_name(), width = 20, anchor = W)
    label2.pack(side = LEFT, fill = X, expand = True)
    label2.bind("<Button-1>", DescriptionWrap(result).fetch_description)
    label3 = Label(row_frame, text = result.get_price(), width = 8, anchor = E)
    label3.pack(side = LEFT)
    label3.bind("<Button-1>", DescriptionWrap(result).fetch_description)
    for label in row_frame.winfo_children():
        label["style"] = "White.TLabel"

def _result_handling(result_list):
    if len(result_list) == 0:
        Label(disposable_frame, text = TAGS["result_no_match"]).pack(fill = X)
    else:
        for result in result_list:
            _insert_result(result)

def _make_all_white():
    for row in disposable_frame.winfo_children():
        for label in row.winfo_children():
            label["style"] = "White.TLabel"

def _make_this_blue(widget):
    parent_name = widget.winfo_parent()
    parent = widget._nametowidget(parent_name)
    for label in parent.winfo_children():
        label["style"] = "Blue.TLabel"

def _insert_description(product, aFrame):
    Label(aFrame, text = TAGS["title_code"]).grid(row = 0, column = 0, sticky = E)
    Label(aFrame, text = product.get_code()).grid(row = 0, column = 1, sticky = W)
    Label(aFrame, text = TAGS["title_name"]).grid(row = 1, column = 0, sticky = E)
    Label(aFrame, text = product.get_name()).grid(row = 1, column = 1, sticky = W)
    Label(aFrame, text = TAGS["title_price"]).grid(row = 2, column = 0, sticky = E)
    Label(aFrame, text = product.get_price()).grid(row = 2, column = 1, sticky = W)
    if _expert_mode:
        Label(aFrame, text = TAGS["title_price_no_vat"]).grid(row = 3, column = 0, sticky = E)
        Label(aFrame, text = product.get_price_no_vat()).grid(row = 3, column = 1, sticky = W)
        Label(aFrame, text = TAGS["title_cost"]).grid(row = 4, column = 0, sticky = E)
        Label(aFrame, text = product.get_cost()).grid(row = 4, column = 1, sticky = W)
        Label(aFrame, text = TAGS["title_gain"]).grid(row = 5, column = 0, sticky = E)
        Label(aFrame, text = product.get_gain()).grid(row = 5, column = 1, sticky = W)
        Label(aFrame, text = TAGS["title_gain_percentage"]).grid(row = 6, column = 0, sticky = E)
        Label(aFrame, text = product.get_gain_percentage()).grid(row = 6, column = 1, sticky = W)
    for label in aFrame.winfo_children():
        label["style"] = "Description.TLabel"

def _make_scrollableter():
    dfooter_frame = Frame(details_frame)
    dfooter_frame.pack(side = BOTTOM, fill = X)
    Button(dfooter_frame, text= TAGS["clear_button"], command = clean_details).pack()

# classes
class ScrollFrame(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)

        self.canvas = Canvas(parent)
        self.myscrollbar = Scrollbar(parent, orient = "vertical", command = self.canvas.yview)
        self.canvas.configure(yscrollcommand = self.myscrollbar.set)
        self.canvas.pack(side = "left", fill = BOTH, expand = True)
        self.myscrollbar.pack(side = "right", fill = "y")

        self.frame = Frame(self.canvas)
        self.win_id = self.canvas.create_window((0, 0), window = self.frame, anchor = NW)
        self.frame.bind("<Configure>", self.region)
        root.bind("<MouseWheel>", self.scroll)
        self.canvas.bind("<Configure>", self.set_width)

    def region(self, event):
        self.canvas.configure(scrollregion = self.canvas.bbox("all"))

    def set_width(self, event):
        self.canvas.itemconfigure(self.win_id, width = event.width)

    def scroll(self, event):
        if self.frame.winfo_height() > self.canvas.winfo_height():
            self.canvas.yview_scroll(-1*(event.delta/120), "units")

class DescriptionWrap:
    def __init__(self, product):
        self.product = product

    def fetch_description(self, event=None):
        clean_details()
        description_frame = Frame(details_frame)
        description_frame.pack(side = TOP)
        widget = event.widget
        if widget["style"] == "White.TLabel":
            _make_all_white()
            _insert_description(self.product, description_frame)
            _make_this_blue(widget)
            _make_scrollableter()
        else:
            _make_all_white()

class WordChangeWindow:
    def __init__(self):
        window = Toplevel()
        window.title(TAGS["title_wordchange"])
        window.iconbitmap(app_icon)
        window.focus()
        self.add_frame = LabelFrame(window, text = TAGS["wordchange_add_frame"])
        self.add_frame.pack(fill = BOTH, expand = True, pady = (0, 15))
        self.delete_frame = LabelFrame(window, text = TAGS["wordchange_delete_frame"])
        self.delete_frame.pack(fill = BOTH, expand = True)
        self.result_frame = LabelFrame(window, text = TAGS["wordchange_result_frame"])
        self.result_frame.pack(fill = BOTH, expand = True, pady = (15, 0))
        self.footer_frame = Frame(window)
        self.footer_frame.pack(fill = X)
        Button(self.footer_frame, text = TAGS["wordchange_showall_button"], command=self.show_all).pack(side = LEFT)
        Button(self.footer_frame, text=TAGS["wordchange_exit_button"], command=window.destroy).pack(side = RIGHT)

        Label(self.add_frame, text = TAGS["wordchange_add_from"]).pack(side = LEFT)
        self.input_entry = Entry(self.add_frame)
        self.input_entry.pack(side = LEFT)
        self.input_entry.focus()
        Label(self.add_frame, text = TAGS["wordchange_add_to"]).pack(side = LEFT)
        self.output_entry = Entry(self.add_frame)
        self.output_entry.pack(side = LEFT)
        Button(self.add_frame, text = TAGS["wordchange_add_button"], command = self.add_wordchange).pack(side = LEFT)
        Label(self.delete_frame, text = TAGS["wordchange_delete"]).pack(side = LEFT)
        self.delete_entry = Entry(self.delete_frame)
        self.delete_entry.pack(side = LEFT)
        Button(self.delete_frame, text = TAGS["wordchange_delete_button"], command = self.delete_wordchange).pack(side = LEFT)

    def add_wordchange(self):
        self.clear_results()
        note = TAGS["wordchange_add_fail_message"]
        if _word_change.add(self.input_entry.get(), self.output_entry.get()):
            note = TAGS["wordchange_success_message"]
            self.result_frame.after(2000, self.clear_results)
        Label(self.result_frame, text = note, style = "Message.TLabel").pack(side = LEFT)
        self.input_entry.delete(0, END)
        self.output_entry.delete(0, END)

    def delete_wordchange(self):
        self.clear_results()
        note = TAGS["wordchange_delete_fail_message"]
        if _word_change.delete(self.delete_entry.get()):
            note = TAGS["wordchange_success_message"]
            self.result_frame.after(2000, self.clear_results)
        Label(self.result_frame, text = note, style = "Message.TLabel").pack(side = LEFT)
        self.delete_entry.delete(0, END)

    def show_all(self):
        self.clear_results()
        tmp = Frame(self.result_frame)
        scroll_frame = ScrollFrame(tmp)
        tmp.pack()
        keys = _word_change.word_change.keys()
        keys.sort()
        for key in keys:
            row_frame = Frame(scroll_frame.frame)
            row_frame.pack(fill = X, expand = True)
            Label(row_frame, text = key, width = 15, style = "White.TLabel").pack(side=LEFT)
            Label(row_frame, text = _word_change.word_change[key], width = 15, style = "White.TLabel").pack(side=LEFT)

    def clear_results(self):
        if len(self.result_frame.winfo_children()) == 1:
            self.result_frame.winfo_children()[0].destroy()

    def __del__(self):
        root.unbind("<MouseWheel>")
        root.bind("<MouseWheel>", scrollable.scroll)

# event handlers
def fetch_results(event=None):
    clean_up()
    result_list = search_by_name(name_entry.get())
    _result_handling(result_list)
    scrollable.canvas.yview_moveto(0)
    name_entry.select_range(0, END)

def fetch_code_results(event=None):
    clean_up()
    result_list = search_by_code(code_entry.get().strip())
    _result_handling(result_list)
    scrollable.canvas.yview_moveto(0)
    code_entry.select_range(0, END)

def clean_up(event=None):
    global disposable_frame
    if len(scrollable.frame.winfo_children())!=0:
        scrollable.frame.winfo_children()[0].destroy()
    disposable_frame = Frame(scrollable.frame)
    disposable_frame.pack(fill = BOTH, expand = True)

def clean_details(event=None):
    global details_frame
    details_frame.destroy()
    details_frame = Frame(root)
    details_frame.pack(side = RIGHT, fill = Y)

def open_wordchange_window(event=None):
    WordChangeWindow()

def toggle_expertmode(event=None):
    global _expert_mode
    _expert_mode = not _expert_mode
    clean_details()

def import_excel(event=None):
    global _all_product_list
    filename = askopenfilename()
    with open(data_location, 'w', 'utf-8') as f:
        f.write(filename)
    del _all_product_list
    _all_product_list = pool(filename, _file_info)

def to_name_tab(event=None):
    notebook.select(notebook.tabs()[0])

def to_code_tab(event=None):
    notebook.select(notebook.tabs()[1])

def open_help(event=None):
    startfile(help_file_location)

# create main window
root = Tk()
root.title(TAGS["title"])
root.iconbitmap(app_icon)

# MAIN
_word_change = WordChange()
with open(data_location) as f:
    data_file = f.readline().decode('utf-8-sig').strip()
with open(data_start) as f:
    _file_info = []
    for line in f:
        _file_info.append(line.decode('utf-8-sig').strip())
_all_product_list = pool(data_file, _file_info)
_expert_mode = False

# style
style = Style()
style.configure("White.TLabel", background = "white", foreground = "black", relief = GROOVE)
style.configure("Blue.TLabel", background = "#0099FF", foreground = "white", relief = GROOVE)
style.configure("Message.TLabel", foreground = "red", justify = LEFT)
style.configure("Description.TLabel", padding = (5, 0, 5, 0))

# menu_bar
menu_bar = Menu(root)

file_menu = Menu(menu_bar, tearoff=0)
file_menu.add_command(label=TAGS["import_command"], command=import_excel)
file_menu.add_separator()
file_menu.add_command(label=TAGS["exit_command"], command=root.destroy)

view_menu = Menu(menu_bar, tearoff=0)
view_menu.add_command(label=TAGS["name_view_command"], command=to_name_tab)
view_menu.add_command(label=TAGS["code_view_command"], command=to_code_tab)
view_menu.add_separator()
view_menu.add_checkbutton(label=TAGS["expert_command"], command=toggle_expertmode)

edit_menu = Menu(menu_bar, tearoff=0)
edit_menu.add_command(label=TAGS["wordchange_command"], command=open_wordchange_window)

help_menu = Menu(menu_bar, tearoff=0)
help_menu.add_command(label=TAGS["help_command"], command=open_help)

menu_bar.add_cascade(label=TAGS["file_menu"], menu=file_menu)
menu_bar.add_cascade(label=TAGS["view_menu"], menu=view_menu)
menu_bar.add_cascade(label=TAGS["edit_menu"], menu=edit_menu)
menu_bar.add_cascade(label=TAGS["help_menu"], menu=help_menu)
root.configure(menu=menu_bar)

# frame layout
main_frame = Frame(root)
main_frame.pack(side = LEFT, fill = BOTH, expand = True)
details_frame = Frame(root)
details_frame.pack(side = RIGHT, fill = Y)

# main_frame
input_frame = Frame(main_frame)
input_frame.pack(fill = X, pady = (0, 10))
output_frame = Frame(main_frame)
output_frame.pack(fill = BOTH, expand = True)

# # input_frame
notebook = Notebook(input_frame)
search_name_frame = Frame(notebook)
search_code_frame = Frame(notebook)
notebook.add(search_name_frame, text = TAGS["search_name_tab"])
notebook.add(search_code_frame, text = TAGS["search_code_tab"])
notebook.pack(fill = X)

# # # search_name_frame
Label(search_name_frame, text = TAGS["search_name"]).pack(side = LEFT, pady = 10, padx = 10)

name_entry = Entry(search_name_frame)
name_entry.pack(side = LEFT, fill = X, expand = True, pady = 10)
name_entry.bind("<Return>", fetch_results)
name_entry.focus()

Button(search_name_frame, text = TAGS["search_button"], command = fetch_results).pack(side = LEFT, pady = 10, padx = 10)

# # # search_code_frame
Label(search_code_frame, text = TAGS["search_code"]).pack(side = LEFT, pady = 10, padx = 10)

code_entry = Entry(search_code_frame)
code_entry.pack(side = LEFT, fill = X, expand = True, pady = 10)
code_entry.bind("<Return>", fetch_code_results)
code_entry.focus()

Button(search_code_frame, text = TAGS["search_button"], command = fetch_code_results).pack(side = LEFT, pady = 10, padx = 10)

# # output_frame
title_frame = Frame(output_frame)
title_frame.pack(fill = X)
Label(title_frame, text = TAGS["title_code"], anchor = W, width = 12).pack(side = LEFT)
Label(title_frame, text = TAGS["title_name"], anchor = W, width = 20).pack(side = LEFT, fill = X, expand = True)
Label(title_frame, text = TAGS["title_price"], anchor = E, width = 8).pack(side = LEFT, padx = (0, 15))

# # # result_frame
result_frame = Frame(output_frame)
scrollable = ScrollFrame(result_frame)
result_frame.pack(fill = BOTH, expand = True)

footer_frame = Frame(output_frame)
footer_frame.pack(fill = X)
Button(footer_frame, text = TAGS["clear_button"], command = clean_up).pack(side = LEFT)

# display
root.mainloop()