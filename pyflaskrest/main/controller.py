from __future__ import print_function
from flask import Blueprint
from flask import request, make_response, json, jsonify, abort
from docx import Document
from mailmerge import MailMerge
from pyflaskrest.config import config
import os.path
import uuid

main = Blueprint('main', __name__)

@main.route('/')
def index():
    return "App is Up!!"

@main.route('/pdfgenerate', methods=['POST'])
def generatepdf():
    if not request.json:
        abort(400)
    root_directory = os.path.dirname(os.path.dirname(__file__))
    proposal_template_path = os.path.join(root_directory, "./templates/SG_Proposal_Full.docx")
    docuuid = uuid.uuid4()
    word_doc_path = os.path.join(root_directory,"./temp/" + str(docuuid) + '.docx')

    # Mail Merge Proposal
    print(proposal_template_path)
    proposal_template_document = MailMerge(proposal_template_path)
    proposal_merge_Fields = proposal_template_document.get_merge_fields()
    merge_field_values = {}
    for field in proposal_merge_Fields:
        merge_field_values[field] = request.json.get(field, "") 
    quotedPlans = request.json.get("QuotedPlans", "")
    
    quotedPlan = quotedPlans[0]
    BusinessPackageId = {'BusinessPackageId' : quotedPlan.get("BusinessPackageId", "")}
    MonthlyPremium = {'MonthlyPremium' : quotedPlan.get("MonthlyPremium", "")}
    SBCs = quotedPlan.get("SBC", "")
    quote_line_census = quotedPlan.get("QuoteCensus", "")
    proposal_template_document.merge(**BusinessPackageId)
    proposal_template_document.merge(**MonthlyPremium)
    proposal_template_document.merge_rows('Name', SBCs)
    proposal_template_document.merge_rows('EmployeeName', quote_line_census)
    proposal_template_document.merge(**merge_field_values)
    proposal_template_document.write(word_doc_path)
    #Convert to PDF
    jar_file_path = os.path.abspath(os.path.join(root_directory, "../bin/docs-to-pdf-converter-1.8.jar"))
    exec_args = " -i " + os.path.abspath(word_doc_path)
    os.system("java -jar " + jar_file_path + exec_args)
    
    pdfresponse = {
        'requestStatus' : 'success' ,
        'documentid' : str(docuuid)
    }
    return jsonify({'response' : pdfresponse}), 201

@main.errorhandler(404)
def not_found(error):
    return make_response(jsonify({'error': 'Resource Not Available'}), 404)