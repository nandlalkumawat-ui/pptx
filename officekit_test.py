from officekit import create_document

output_file = "test_officekit.pptx"

create_document(
    "pptx",
    output_file,
    title="OfficeKit Test Presentation"
)

print(f"Created: {output_file}")