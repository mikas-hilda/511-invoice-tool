import io
import zipfile
import tempfile
from datetime import datetime
from pathlib import Path

import streamlit as st

from converter_511_invoice import convert_invoice_pdf, InvoiceConversionError


st.set_page_config(page_title="5.11 Invoice PDF → Excel", layout="centered")

st.title("5.11 Invoice PDF → Excel")
st.write(
    "Įkelk vieną arba kelias 5.11 PDF sąskaitas. "
    "Įrankis sugeneruos Excel tik tiems failams, kurie praeis validaciją."
)

uploaded_files = st.file_uploader(
    "PDF sąskaitos",
    type=["pdf"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.write(f"Įkelta PDF failų: **{len(uploaded_files)}**")

    successful_files = []
    failed_files = []
    validation_rows = []

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        for uploaded in uploaded_files:
            pdf_path = tmpdir / uploaded.name
            xlsx_path = tmpdir / (Path(uploaded.name).stem + ".xlsx")

            pdf_path.write_bytes(uploaded.read())

            try:
                result = convert_invoice_pdf(pdf_path, xlsx_path)
                xlsx_bytes = xlsx_path.read_bytes()

                successful_files.append(
                    {
                        "pdf_name": uploaded.name,
                        "xlsx_name": xlsx_path.name,
                        "xlsx_bytes": xlsx_bytes,
                        "result": result,
                    }
                )

                validation_rows.append(
                    {
                        "PDF": uploaded.name,
                        "Status": "OK",
                        "Invoice Number": result.invoice.invoice_number,
                        "Invoice Date": result.invoice.invoice_date,
                        "Customer PO": result.invoice.customer_po,
                        "Invoice Total": result.invoice.invoice_total,
                        "PDF Rows": result.pdf_row_count,
                        "Excel Rows": result.excel_row_count,
                        "Qty Sum": result.qty_sum,
                        "Amount Sum": result.amount_sum,
                        "Error": "",
                    }
                )

            except InvoiceConversionError as exc:
                failed_files.append({"pdf_name": uploaded.name, "error": str(exc)})
                validation_rows.append(
                    {
                        "PDF": uploaded.name,
                        "Status": "ERROR",
                        "Invoice Number": "",
                        "Invoice Date": "",
                        "Customer PO": "",
                        "Invoice Total": "",
                        "PDF Rows": "",
                        "Excel Rows": "",
                        "Qty Sum": "",
                        "Amount Sum": "",
                        "Error": str(exc),
                    }
                )

            except Exception as exc:
                failed_files.append({"pdf_name": uploaded.name, "error": repr(exc)})
                validation_rows.append(
                    {
                        "PDF": uploaded.name,
                        "Status": "ERROR",
                        "Invoice Number": "",
                        "Invoice Date": "",
                        "Customer PO": "",
                        "Invoice Total": "",
                        "PDF Rows": "",
                        "Excel Rows": "",
                        "Qty Sum": "",
                        "Amount Sum": "",
                        "Error": repr(exc),
                    }
                )

        if successful_files:
            st.success(f"Sėkmingai konvertuota: {len(successful_files)} PDF")

            st.write("### Sėkmingai sugeneruoti failai")

            for item in successful_files:
                result = item["result"]

                with st.expander(f"{item['pdf_name']} → {item['xlsx_name']}", expanded=False):
                    st.write("**Kontrolė:**")
                    for validation in result.validations:
                        st.write(f"- {validation}")

                    st.write("**Antraštė:**")
                    st.write(f"- Invoice Number: `{result.invoice.invoice_number}`")
                    st.write(f"- Invoice Date: `{result.invoice.invoice_date}`")
                    st.write(f"- Customer PO: `{result.invoice.customer_po}`")
                    st.write(f"- Invoice Total: `{result.invoice.invoice_total}`")

                    st.download_button(
                        label=f"Atsisiųsti {item['xlsx_name']}",
                        data=item["xlsx_bytes"],
                        file_name=item["xlsx_name"],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_{item['xlsx_name']}",
                    )

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for item in successful_files:
                    zf.writestr(item["xlsx_name"], item["xlsx_bytes"])

                report_lines = [
                    "5.11 Invoice PDF -> Excel validation report",
                    f"Generated: {datetime.now().isoformat(timespec='seconds')}",
                    "",
                ]

                for row in validation_rows:
                    report_lines.append(f"PDF: {row['PDF']}")
                    report_lines.append(f"Status: {row['Status']}")

                    if row["Status"] == "OK":
                        report_lines.append(f"Invoice Number: {row['Invoice Number']}")
                        report_lines.append(f"Invoice Date: {row['Invoice Date']}")
                        report_lines.append(f"Customer PO: {row['Customer PO']}")
                        report_lines.append(f"Invoice Total: {row['Invoice Total']}")
                        report_lines.append(f"PDF Rows: {row['PDF Rows']}")
                        report_lines.append(f"Excel Rows: {row['Excel Rows']}")
                        report_lines.append(f"Qty Sum: {row['Qty Sum']}")
                        report_lines.append(f"Amount Sum: {row['Amount Sum']}")
                    else:
                        report_lines.append(f"Error: {row['Error']}")

                    report_lines.append("-" * 60)

                zf.writestr("validation_report.txt", "\n".join(report_lines))

            zip_buffer.seek(0)

            st.download_button(
                label="Atsisiųsti visus Excel kaip ZIP",
                data=zip_buffer.getvalue(),
                file_name="converted_invoices.zip",
                mime="application/zip",
                key="download_all_zip",
            )

        if failed_files:
            st.error(f"Nepavyko konvertuoti: {len(failed_files)} PDF")

            st.write("### Klaidos")

            error_report_lines = []
            for item in failed_files:
                st.write(f"**{item['pdf_name']}**")
                st.code(item["error"])
                error_report_lines.append(f"PDF: {item['pdf_name']}")
                error_report_lines.append(f"Error: {item['error']}")
                error_report_lines.append("-" * 60)

            st.download_button(
                label="Atsisiųsti klaidų logą",
                data="\n".join(error_report_lines).encode("utf-8"),
                file_name="conversion_errors.txt",
                mime="text/plain",
                key="download_errors",
            )

        st.write("### Bendras rezultatas")
        st.dataframe(validation_rows, use_container_width=True)
