import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import tempfile

EXPECTED_COLUMNS = [
    "Name", "Identifier", "FirstName", "MiddleName", "LastName",
    "PlanCost", "EmploymentStatus", "HireDate", "HiredOn",
    "TerminationDate", "TerminatedOn", "StartDate",
    "EnrolledOn", "EndDate", "EndedOn" , "CoverageLevel"
]

def show_xml_converter():
    def xml_to_exact_excel(xml_file):
        # Parse XML
        context = ET.iterparse(xml_file, events=("start", "end"))
        data_list = []
        header_info = {}

        for event, elem in context:
            if event == "end":
                if elem.tag == "Header":
                    header_info = {
                        "Disclaimer": elem.findtext("Disclaimer", ""),
                        "ExchangeName": elem.findtext("ExchangeName", ""),
                        "VendorName": elem.findtext("VendorName", ""),
                        "RunDate": elem.findtext("RunDate", ""),
                    }

                elif elem.tag == "Company":
                    company_data = {
                        "Identifier": elem.findtext("Identifier", ""),
                        "Name": elem.findtext("Name", ""),
                    }
                    full_data = {**header_info, **company_data}
                    employees = elem.findall("Employees/Employee")
                    first_employee = True
                    
                    for emp in employees:
                        emp_data = {
                            "FirstName": emp.findtext("FirstName", ""),
                            "MiddleName": emp.findtext("MiddleName", ""),
                            "LastName": emp.findtext("LastName", ""),
                            "EmploymentStatus": emp.findtext("EmploymentStatus", ""),
                            "HireDate": emp.findtext("HireDate", ""),
                            "HiredOn": emp.findtext("HiredOn", ""),
                            "TerminationDate": emp.findtext("TerminationDate", ""),
                            "TerminatedOn": emp.findtext("TerminatedOn", ""),
                        }
                        
                        enrollments = emp.findall("Enrollments/Enrollment")
                        if enrollments:
                            for enroll in enrollments:
                                enroll_data = {
                                    "PlanCost": enroll.findtext("PlanCost", ""),
                                    "StartDate": enroll.findtext("StartDate", ""),
                                    "EnrolledOn": enroll.findtext("EnrolledOn", ""),
                                    "EndDate": enroll.findtext("EndDate", ""),
                                    "EndedOn": enroll.findtext("EndedOn", ""),
                                    "CoverageLevel": enroll.findtext("CoverageLevel", "")
                                }
                                
                                if first_employee:
                                    final_data = {**full_data, **emp_data, **enroll_data}
                                    first_employee = False
                                else:
                                    final_data = {**{k: "" for k in full_data}, **emp_data, **enroll_data}
                                
                                cleaned_data = {col: final_data.get(col, "") for col in EXPECTED_COLUMNS}
                                data_list.append(cleaned_data)
                        else:
                            if first_employee:
                                final_data = {**full_data, **emp_data}
                                first_employee = False
                            else:
                                final_data = {**{k: "" for k in full_data}, **emp_data}
                            
                            cleaned_data = {col: final_data.get(col, "") for col in EXPECTED_COLUMNS}
                            data_list.append(cleaned_data)
                    elem.clear()
        
        df = pd.DataFrame(data_list, columns=EXPECTED_COLUMNS)
        return df

    st.title("XML to Excel Converter")

    uploaded_xml = st.file_uploader("Upload XML File", type=["xml"])

    if uploaded_xml:
        with tempfile.NamedTemporaryFile(delete=False) as tmp_xml:
            tmp_xml.write(uploaded_xml.read())
            tmp_xml.close()
            
            df = xml_to_exact_excel(tmp_xml.name)
            st.write("Preview of Extracted Data:")
            st.dataframe(df.head())
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_out:
                df.to_excel(tmp_out.name, index=False)
                st.download_button(
                    label="XML Converted! Download Excel",
                    data=open(tmp_out.name, "rb").read(),
                    file_name="converted.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )