from excel_handler import ensure_excel, append_row, read_all_rows, email_duplicate_within_days, save_to_excel
sample = {
    "Name": "Test User",
    "Email": "test.user@example.com",
    "Phone": "+91 9876543210",
    "Skills": "Python, Excel",
    "Experience": "3 years",
    "ResumePath": "sample_resume.pdf"
}
ensure_excel("test_output.xlsx")
save_to_excel(sample, "test_output.xlsx")
print("Rows now:", read_all_rows("test_output.xlsx"))
print("Duplicate check (should be True):", email_duplicate_within_days("test_output.xlsx", "test.user@example.com", 365))
