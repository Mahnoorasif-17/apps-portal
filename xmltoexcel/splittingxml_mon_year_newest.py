import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from io import BytesIO

def show_xml_converter():
    st.header("📑 XML to Excel Converter")
    st.write("Upload your Redirect Health XML file to convert it into a standardized Excel format.")

    # --- 1. EXPECTED COLUMNS ---
    EXPECTED_COLUMNS = [
        "Name", "Identifier", "FirstName", "MiddleName", "LastName",
        "PlanCost", "EmploymentStatus", "HireDate", "HiredOn",
        "TerminationDate", "TerminatedOn", "StartDate",
        "EnrolledOn", "EndDate", "EndedOn", "CoverageLevel",
        "CarrierPlanCode", "PriorCoverageStartDate"
    ]

    # --- 2. FILE UPLOADER ---
    uploaded_file = st.file_uploader("Choose an XML file", type="xml")

    if uploaded_file is not None:
        if st.button("Convert XML to Excel"):
            try:
                # Use iterparse on the uploaded file stream
                context = ET.iterparse(uploaded_file, events=("start", "end"))
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
                                            "CoverageLevel": enroll.findtext("CoverageLevel", ""),
                                            "CarrierPlanCode": enroll.findtext("CarrierPlanCode", ""),
                                            "PriorCoverageStartDate": enroll.findtext("PriorCoverageStartDate", ""),
                                        }

                                        if first_employee:
                                            final_data = {**full_data, **emp_data, **enroll_data}
                                            first_employee = False
                                        else:
                                            # Keep company/header info empty for rows after the first for a clean look
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

                # --- 3. CREATE DATAFRAME & EXPORT ---
                df = pd.DataFrame(data_list, columns=EXPECTED_COLUMNS)
                
                # Convert DF to Excel in Memory
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                
                processed_data = output.getvalue()

                st.success("✅ Conversion Successful!")
                
                # --- 4. DOWNLOAD BUTTON ---
                st.download_button(
                    label="📥 Download Excel File",
                    data=processed_data,
                    file_name="Converted_XML_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"An error occurred: {e}")