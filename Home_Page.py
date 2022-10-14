import time

from pathlib import Path
import streamlit as st
import pandas as pd
import numpy as np
from streamlit_option_menu import option_menu
from st_aggrid import AgGrid
from st_aggrid import GridOptionsBuilder
from st_aggrid import GridUpdateMode
from streamlit_autorefresh import st_autorefresh

st.title('Product Installation automation')

## Assignation fichier XSAMP, TESTSAMP & KARAP
productsfile = "products2.xlsx"
sheet_name1 = 'Sheet1'
usecols1 = 'A:D'

waferpnsfile = 'waferpns.xlsx'
sheet_name2 = 'Sheet1'
usecols2 = 'A:J'


def progress_bar():
    latest_status = st.empty()
    bar_progress = st.progress(0)
    for i in range(100):
        latest_status.text(f"En cours ... {i+1}")
        bar_progress.progress(i+1)
        time.sleep(0.01)
        st.empty()

def upload_excel_file(file, sheet_name, usecols):
    df_products = pd.read_excel(file,
                                engine='openpyxl',
                                sheet_name=sheet_name,
                                usecols=usecols,
                                header=0)
    return df_products

# Horizontal main menu
selection = option_menu(
    menu_title=None,
    options=["XsampBak", "TestSampBak", "Karap"],
    #icons="",
    #menu_icon="",
    default_index=0,
    orientation="horizontal",
)

def aggrid_interactive(df0):
    ## AG grid for Updating files and interactive dashboard
    gd = GridOptionsBuilder.from_dataframe(df0)
    gd.configure_pagination(enabled=True)
    gd.configure_default_column(editable=True, groupable=True)

    sel_option = st.radio('Les fonctionalités', options= ['Lecuture', 'Update'])
    gd.configure_selection(selection_mode=sel_option, use_checkbox=True)
    grid_options = gd.build()
    grid_table = AgGrid(productsfile,
                        grid_Options=grid_options,
                        update_mode=GridUpdateMode.SELECTION_CHANGED,
                        #columns_auto_size_mode=,
       allow_unsafe_jscode=True,
       theme='balham')

def update_form(df0, file):
    data_folder = Path("ProductInstallation/")
    print(f"Folder Path =", data_folder)
    file_to_open = data_folder / file
    print(f"File path = ", file_to_open)

    ##For Second version of APP
    def xsampbak_form():
        ## Options for XSAMBAK
        options_form = st.sidebar.form("Remplisser les cases pour update")
        Kalist = options_form.number_input("Kalist")
        Description = options_form.text_input("Description")
        Extras = options_form.text_input("Extras")
        Options = options_form.text_input("Options")
        ajout_update = options_form.form_submit_button()

    def testsamp_form():
        options_form = st.sidebar.form("Remplisser les cases pour update")
        ## Options for TESTSAMPBAK
        Waferpn = options_form.text_input("Waferpn")
        Stepplan = options_form.text_input("Stepplan")
        StepplanOverride = options_form.text_input("StepplanOverride")
        DoublePass = options_form.text_input("DoublePass")
        WLR = options_form.text_input("WLR")
        Version = options_form.text_input("Version")
        Comments = options_form.text_input("Comments")
        ajout_update = options_form.form_submit_button()

    file = None
    if selection == "XsampBak":

            ## Options for XSAMBAK
            options_form = st.sidebar.form("Remplisser les cases pour update Xsampbak")
            Kalist = options_form.number_input("Kalist")
            Description = options_form.text_input("Description")
            Extras = options_form.text_input("Extras")
            Options = options_form.text_input("Options")
            ajout_update = options_form.form_submit_button()

            if ajout_update:
                new_data = {"Kalist": Kalist, "Description": Description, "Extras": Extras, "Options": Options}
                df0 = df0.append(new_data, ignore_index=True)
                try:
                    print(file_to_open)
                    df0.to_excel("products2.xlsx", index=False)
                    #update_file()
                except:
                    print("There is no file dude")

    elif selection == "TestSampBak":
            options_form = st.sidebar.form("Remplisser les cases pour update TestsampBak")
            ## Options for TESTSAMPBAK
            Waferpn = options_form.text_input("Waferpn")
            Stepplan = options_form.text_input("Stepplan")
            Kalist = options_form.number_input("Kalist")
            StepplanOverride = options_form.text_input("StepplanOverride")
            DoublePass = options_form.text_input("DoublePass")
            WLR = options_form.text_input("WLR")
            Version = options_form.number_input("Version")
            Comments = options_form.text_input("Comments")
            Product_Code = options_form.text_input("Product Code")
            Family_Code = options_form.text_input("Family Code")
            ajout_update2 = options_form.form_submit_button()

            if ajout_update2:
                new_data = {"Waferpn": Waferpn, "Stepplan": Stepplan, "Kalist": Kalist, "StepplanOverride": StepplanOverride, "DoublePass": DoublePass, "WLR": WLR, "Version": Version, "Comments": Comments, "Product Code": Product_Code, "Family Code": Family_Code}
                df0 = df0.append(new_data, ignore_index=True)
                try:
                    print(file_to_open)
                    df0.to_excel("waferpns2.xlsx")
                    #update_file()
                except:
                    print("There is no file dude")

# For refreshing Page
def update_file():
    st_autorefresh(interval=20, limit=1, key="Refresh")

#Check print data frame
if selection == "XsampBak":
    progress_bar()
    df_product = upload_excel_file(productsfile, sheet_name1, usecols1)
    #df0 = df_product
    st.subheader("DATAFRAME PRODUCTS")
    st.dataframe(df_product)
    #AgGrid(productsfile)
    #aggrid_interactive(df_product)
    st.info(f"Le nombre de produits installé est: {len(df_product)}")
    update_form(df_product, productsfile)
    #update_file()

elif selection == "TestSampBak":
    progress_bar()
    st.subheader("DATAFRAME WFERPNS")
    df_waferpns = upload_excel_file(waferpnsfile, sheet_name2, usecols2)
    st.dataframe(df_waferpns)
    st.info(f"Le nombre de waferpns installé est: {len(df_waferpns)}")
    update_form(df_waferpns, waferpnsfile)
    st.dataframe(df_waferpns)

#elif selection == "Karap":
#    progress_bar()
#    st.subheader('DATAFRAME KARAP')
#    df_karap = upload_excel_file(karap, sheet_name3, usecols3)
#    st.dataframe(df_karap)

