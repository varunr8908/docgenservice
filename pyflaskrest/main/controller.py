from __future__ import print_function
from flask import Blueprint
from flask import request, make_response, json, jsonify, abort
from docx import Document
from mailmerge import MailMerge
from pyflaskrest.config import config
import os.path
import platform
if platform.system() == 'Windows':
    import comtypes
    import comtypes.client
import uuid
import base64

main = Blueprint('main', __name__)


@main.route('/')
def index():
    return "App is Up!!"

@main.route('/pdfgenerate', methods=['POST'])
def generatepdf():
    if not request.json:
        abort(400)
    root_directory = os.path.dirname(os.path.dirname(__file__))
    #Fetch Required Attributes From Request
    doc_format = request.json.get("DocFormat")
    cover_template_encoded = str(request.json.get("CoverTemplate"))
    footer_template_encoded = str(request.json.get("FooterTemplate"))
    
    platform_name = platform.system()
    #Generate a GUID For the Transaction
    docuuid = uuid.uuid4()

    #Set Template Paths
    merged_doc_path = os.path.join(root_directory,"./temp/" + str(docuuid) + '.docx')
    merged_pdf_path =  os.path.join(root_directory,"./temp/" + str(docuuid) + '.pdf')
    cover_template_path = os.path.join(root_directory,"./temp/" + str(docuuid) + '_covertemplate.docx')
    cover_page_path = os.path.join(root_directory,"./temp/" + str(docuuid) + '_coverpage.docx')
    footer_template_path = os.path.join(root_directory,"./temp/" + str(docuuid) + '_footertemplate.docx')
    footer_page_path = os.path.join(root_directory,"./temp/" + str(docuuid) + '_footerpage.docx')
    input_files = [cover_page_path]
    #Write To Templates
    cover_template_decoded = base64.b64decode(cover_template_encoded)
    fh = open(cover_template_path, "wb")
    fh.write(cover_template_decoded)
    fh.close()

    footer_template_decoded = base64.b64decode(footer_template_encoded)
    fh = open(footer_template_path, "wb")
    fh.write(footer_template_decoded)
    fh.close()

    #Mail Merge Cover Letter and Footer
    cover_letter_template = MailMerge(cover_template_path)
    cover_letter_merge_fields = cover_letter_template.get_merge_fields()
    merge_field_values = {}
    for field in cover_letter_merge_fields:
        merge_field_values[field] = request.json.get(field, "")
    cover_letter_template.merge(**merge_field_values)
    cover_letter_template.write(cover_page_path)

    footer_letter_template = MailMerge(footer_template_path)
    footer_letter_merge_fields = footer_letter_template.get_merge_fields()
    merge_field_values = {}
    for field in footer_letter_merge_fields:
        merge_field_values[field] = request.json.get(field, "")
    footer_letter_template.merge(**merge_field_values)
    footer_letter_template.write(footer_page_path)

    quotedPlans = request.json.get("QuotedPlans", "")
    for i in range(len(quotedPlans)):
        quotedPlan = quotedPlans[i]
        
        sbc_template_encoded = str(quotedPlan.get("SBCTemplate"))
        sbc_template_path = os.path.join(root_directory,"./temp/" + str(docuuid) + '_' + str(i) + '_sbctemplate.docx')
        sbc_page_path = os.path.join(root_directory,"./temp/" + str(docuuid) + '_' + str(i) + '_sbcpage.docx')
        input_files.append(sbc_page_path)
        #Write to SBC Templates
        sbc_template_decoded = base64.b64decode(sbc_template_encoded)
        fh = open(sbc_template_path, "wb")
        fh.write(sbc_template_decoded)
        fh.close()
        #Mail Merge SBC Template
        sbc_letter_template = MailMerge(sbc_template_path)
        sbc_letter_merge_fields = sbc_letter_template.get_merge_fields()
        for field in sbc_letter_merge_fields:
            if(quotedPlan.get(field, "") != ""):
                merge_field_values[field] = quotedPlan.get(field, "")
        for field in sbc_letter_merge_fields:
            if(request.json.get(field, "") != ""):
                merge_field_values[field] = request.json.get(field, "")
        
        SBCs = quotedPlan.get("SBC", "")
        quote_line_census = quotedPlan.get("QuoteCensus", "")
        sbc_letter_template.merge_rows("Name", SBCs)
        sbc_letter_template.merge_rows("EmployeeName", quote_line_census)
        sbc_letter_template.merge(**merge_field_values)
        sbc_letter_template.write(sbc_page_path)

    #Append Footer
    input_files.append(footer_page_path)
    #Merge the Documents
    merged_document = combine_word_documents(input_files)
    merged_document.save(merged_doc_path)

    merged_doc_encoded = ""

    if(doc_format == 'pdf'):
        #Convert to PDF
        if(platform_name == 'Windows'):
            print('Using Com')
            wdFormatPDF = 17
            comtypes.CoInitialize()
            word = comtypes.client.CreateObject('Word.Application')
            
            doc = word.Documents.Open(merged_doc_path)
            #word.Documents.Merge()
            doc.SaveAs(os.path.abspath(merged_pdf_path), FileFormat=wdFormatPDF)
            doc.Close()
            word.Quit()
        else:
            jar_file_path = os.path.abspath(os.path.join(root_directory, "../bin/docs-to-pdf-converter-1.8.jar"))
            exec_args = " -i " + os.path.abspath(merged_doc_path)
            os.system("java -jar " + jar_file_path + exec_args)  
        
        with open(merged_pdf_path, "rb") as pdf_file:
            merged_doc_encoded = base64.b64encode(pdf_file.read())
    else:
         with open(merged_doc_path, "rb") as docfile:
            merged_doc_encoded = base64.b64encode(docfile.read())
    
    pdfresponse = {
        'requestStatus' : 'success' ,
        'documentid' : str(docuuid) , 
        'documentname' : str(docuuid) + '.' + doc_format,
        'document' : merged_doc_encoded.decode("ascii")
    }
    return jsonify(pdfresponse), 200

@main.errorhandler(404)
def not_found(error):
    return make_response(jsonify({'error': 'Resource Not Available'}), 404)

def encode_as_base64( file_path ):
    encoded_string = ""
    with open(file_path, "rb") as template_file:
        encoded_string = base64.b64encode(template_file.read())

    return encoded_string

def combine_word_documents(input_files):
    """
    :param input_files: an iterable with full paths to docs
    :return: a Document object with the merged files
    """
    for filnr, file in enumerate(input_files):
        # in my case the docx templates are in a FileField of Django, add the MEDIA_ROOT, discard the next 2 lines if not appropriate for you. 
        if filnr == 0:
            merged_document = Document(file)
            

        else:
            sub_doc = Document(file)
            # Don't add a page break if you've reached the last file.
            if filnr < len(input_files)-1:
               sub_doc.add_page_break()
            
            for element in sub_doc.element.body:
                merged_document.element.body.append(element)


    return merged_document
