import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
import os


def generate_report_cards(file_path):
    try:
        # Read data from Excel file
        data = pd.read_excel(file_path)

        # Validate columns
        required_columns = {"Student ID", "Name", "Subject", "Score"}
        if not required_columns.issubset(data.columns):
            raise ValueError("Excel file does not contain the required columns.")

        # Handle missing or invalid data
        data.dropna(subset=["Student ID", "Name", "Subject", "Score"], inplace=True)
        data["Score"] = pd.to_numeric(data["Score"], errors="coerce")
        data.dropna(subset=["Score"], inplace=True)

        # Group data by student
        grouped = data.groupby("Student ID")

        # Generate PDF report card for each student
        for student_id, group in grouped:
            student_name = group["Name"].iloc[0]
            total_score = group["Score"].sum()
            avg_score = group["Score"].mean()

            # Prepare data for the table
            subject_scores = group[["Subject", "Score"]].values.tolist()
            subject_scores.insert(0, ["Subject", "Score"])  # Add header

            # Create PDF
            pdf_name = f"report_card_{student_id}.pdf"
            c = canvas.Canvas(pdf_name, pagesize=letter)
            c.setFont("Helvetica-Bold", 14)
            c.drawString(100, 750, f"Report Card for {student_name}")
            c.setFont("Helvetica", 12)
            c.drawString(100, 720, f"Student ID: {student_id}")
            c.drawString(100, 700, f"Total Score: {total_score}")
            c.drawString(100, 680, f"Average Score: {avg_score:.2f}")

            # Add table for subject-wise scores
            table = Table(subject_scores)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ]))
            table.wrapOn(c, 400, 400)
            table.drawOn(c, 100, 600 - len(subject_scores) * 20)

            # Save the PDF
            c.save()
            print(f"Generated: {pdf_name}")

    except FileNotFoundError:
        print("Error: The specified Excel file was not found.")
    except ValueError as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")


# Specify the path to the Excel file
file_path = "student_scores.xlsx"
generate_report_cards(file_path)
