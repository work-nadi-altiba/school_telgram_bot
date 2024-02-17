# settings.py
from decouple import config

FILL_OFFICIAL_MARKS_A3_TWO_FACE_DOC2_OFFLINE_VERSION_URL_1 = config('fill_official_marks_a3_two_face_doc2_offline_version_url_1')
FILL_OFFICIAL_MARKS_A3_TWO_FACE_DOC2_OFFLINE_VERSION_URL_2 = config('fill_official_marks_a3_two_face_doc2_offline_version_url_2')
ASSESSMENTS_PERIODS_MIN_MAX_MARK_URL = config('assessments_periods_min_max_mark_url')
GET_CLASS_STUDENTS_IDS_URL = config('get_class_students_ids_url')
GET_GRADE_ID_FROM_ASSESSMENT_ID_URL = config('get_grade_id_from_assessment_id_url')
CREATE_E_SIDE_MARKS_DOC_URL = config('create_e_side_marks_doc_url')
FILL_OFFICIAL_MARKS_DOC_WRAPPER_OFFLINE_URL_1 = config('fill_official_marks_doc_wrapper_offline_url_1')
FILL_OFFICIAL_MARKS_DOC_WRAPPER_OFFLINE_URL_2 = config('fill_official_marks_doc_wrapper_offline_url_2')
GET_STUDENTS_MARKS_URL = config('get_students_marks_url')
GET_ASSESSMENTS_PERIODS_URL_1 = config('get_assessments_periods_url_1')
GET_ACADEMIC_TERMS_URL = config('get_AcademicTerms_url')
GET_BASIC_INFO_URL_1 = config('get_basic_info_url_1')
GET_BASIC_INFO_URL_2 = config('get_basic_info_url_2')
GET_AUTH_URL_1 = config('get_auth_url_1')
INST_NAME_URL = config('inst_name_url')
INST_AREA_URL = config('inst_area_url')
USER_INFO_URL = config('user_info_url')
GET_TEACHER_CLASSES1_URL = config('get_teacher_classes1_url')
GET_TEACHER_CLASSES2_URL = config('get_teacher_classes2_url')
GET_CLASS_STUDENTS_URL = config('get_class_students_url')
ENTER_MARK_URL = config('enter_mark_url')
ENTER_MARK_JSON_DATA=config('enter_mark_json_data')
GET_CURR_PERIOD_URL = config('get_curr_period_url')
GET_ASSESSMENTS_URL = config('get_assessments_url')
GET_SUB_INFO_URL = config('get_sub_info_url')
CREATE_EXCEL_SHEETS_MARKS_URL = config('create_excel_sheets_marks_url')
GET_STUDENTS_MARKS_URL = config('get_students_marks_url')
def ENTER_MARK_JSON_DATA_FUN(marks, assessment_grading_option_id, assessment_id, education_subject_id, education_grade_id, institution_id, academic_period_id, institution_classes_id, student_status_id, student_id, assessment_period_id):
    """_summary_

    Args:
        marks (_type_): _description_
        assessment_grading_option_id (_type_): _description_
        assessment_id (_type_): _description_
        education_subject_id (_type_): _description_
        education_grade_id (_type_): _description_
        institution_id (_type_): _description_
        academic_period_id (_type_): _description_
        institution_classes_id (_type_): _description_
        student_status_id (_type_): _description_
        student_id (_type_): _description_
        assessment_period_id (_type_): _description_

    Returns:
        _type_: _description_
    """    
    return {
        'marks': marks,
        'assessment_grading_option_id': assessment_grading_option_id,
        'assessment_id': assessment_id,
        'education_subject_id': education_subject_id,
        'education_grade_id': education_grade_id,
        'institution_id': institution_id,
        'academic_period_id': academic_period_id,
        'institution_classes_id': institution_classes_id,
        'student_status_id': student_status_id,
        'student_id': student_id,
        'assessment_period_id': assessment_period_id,
        'action_type': 'default',
    }
