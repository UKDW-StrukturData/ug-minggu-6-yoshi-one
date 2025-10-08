import streamlit as st
from DatabaseManager import excelManager

em = excelManager("dataExcel.xlsx")
options = ["Choose Action", "Insert", "Edit", "Delete"]
choice = st.selectbox("Choose an action:", options)

# Validasi angka
def is_valid_number(value):
    return value.isdigit()

# DELETE
if choice == "Delete":
    nim = st.text_input("Enter targeted NIM:", key="targetNim")
    saveChange = st.checkbox("SaveChanges", value=False, key="saveChangeEditDelete")
    if st.button("Delete"):
        if not is_valid_number(nim):
            st.error("Input NIM harus berupa angka semua")
        elif not em.getData("NIM", nim):
            st.error("NIM tidak ditemukan")
        else:
            em.deleteData(nim, saveChange)
            if not em.getData("NIM", nim):
                st.success("Data Sukses di Hapus")

# INSERT & EDIT
if choice in ("Insert", "Edit"):
    newNim = st.text_input("Enter New NIM:", key="newNim")
    newName = st.text_input("Enter New Name:", key="newName")
    newGrade = st.text_input("Enter New Grade :", key="newGrade")
    saveChange = st.checkbox("SaveChanges", value=False, key="saveChangeInsertEdit")

    if st.button(choice):
        if not is_valid_number(newNim):
            st.error("Input NIM harus berupa angka semua")
        elif not is_valid_number(newGrade):
            st.error("Input nilai harus berupa angka semua")
        else:
            if choice == "Insert":
                if em.getData("NIM", newNim):
                    st.error("NIM sudah ada")
                else:
                    em.insertData({
                        "NIM": newNim.strip(),
                        "Nama": newName.strip(),
                        "Nilai": int(newGrade.strip())
                    }, saveChange)
                    st.success("Data Sukses di Masukan")
            elif choice == "Edit":
                targetNim = st.text_input("Enter targeted NIM:", key="targetNimEdit")
                if not em.getData("NIM", targetNim):
                    st.error("NIM tidak ditemukan")
                else:
                    em.editData(targetNim, {
                        "NIM": newNim.strip(),
                        "Nama": newName.strip(),
                        "Nilai": int(newGrade.strip())
                    }, saveChange)
                    st.success("Data Sukses di Edit")

# FILTER TABEL
option = ["Default", ">", "<", "=", "<=", ">="]
filterSelectBox = st.selectbox("Sort table by: ", option)

if filterSelectBox == "Default":
    st.table(em.getDataFrame())
else:
    targetFilterColumn = st.selectbox("Target Column", ["NIM", "Nilai"])
    filter_val = st.text_input("Filter Nilai")

    if filter_val:
        if not is_valid_number(filter_val):
            st.error("Filter harus berupa angka")
        else:
            df = em.getDataFrame()
            val = int(filter_val)
            if filterSelectBox == ">":
                st.table(df[df[targetFilterColumn] > val])
            elif filterSelectBox == "<":
                st.table(df[df[targetFilterColumn] < val])
            elif filterSelectBox == "=":
                st.table(df[df[targetFilterColumn] == val])
            elif filterSelectBox == "<=":
                st.table(df[df[targetFilterColumn] <= val])
            elif filterSelectBox == ">=":
                st.table(df[df[targetFilterColumn] >= val])
