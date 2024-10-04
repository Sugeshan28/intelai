import PyPDF2
from transformers import AutoTokenizer, AutoModelForSeq2SeqLM

class Pdftoppt:
    def extract_text_from_pdf(pdf_file_path):
        text = ''
        with open(pdf_file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                text += page.extract_text() + '\n'
        return text.strip()

    def separate_title_and_content(text):
        lines = text.splitlines()
        title = lines[0].strip() if lines else ''
        content = '\n'.join(line.strip() for line in lines[1:] if line.strip()) if len(lines) > 1 else ''
        return title, content

    def summarize_text(input_text):
        tokenizer = AutoTokenizer.from_pretrained("Falconsai/text_summarization")
        model = AutoModelForSeq2SeqLM.from_pretrained("Falconsai/text_summarization")

        # Tokenize the input text
        inputs = tokenizer(input_text, return_tensors="pt", max_length=1024, truncation=True)

        # Generate summary
        summary_ids = model.generate(inputs["input_ids"], max_length=200, num_beams=4, early_stopping=True)

        # Decode the summary
        summary = tokenizer.decode(summary_ids[0], skip_special_tokens=True)
        return summary

    # Example usage
    pdf_path = r'D:\Code\Tot_PPT\static\pdf\mf.pdf'  # Replace with your PDF file path
    extracted_text = extract_text_from_pdf(pdf_path)

    if extracted_text:
        title, content = separate_title_and_content(extracted_text)
        summary = summarize_text(content)
        
        print("Title:")
        print(title)
        print("\nContent Summary:")
        print(summary)
    else:
        print("No text extracted from the PDF.")
