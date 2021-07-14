import datetime
from datetime import datetime as dt


def index():
    return dict(message="Hello!")


def suffix(d):
    return 'th' if 11 <= d <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(d % 10,
                                                                      'th')


def custom_strftime(format_date, t):
    return t.strftime(format_date).replace('{S}', str(t.day) + suffix(t.day))


def page1():
    header = db(db.campus_college_institute.cci_id == 11).select(
        db.campus_college_institute.cci_name,
        db.campus_college_institute.address,
        db.campus_college_institute.tel_no)

    student = db(db.grad_stud.student_id == 1).select(
        db.grad_stud.first_name,
        db.grad_stud.middle_name,
        db.grad_stud.last_name,
        db.grad_stud.gender,
        db.degree_courses.course_name,
        db.degree_courses.course_abbrev,
        db.majors.major,
        left=[db.degree_courses.on(db.grad_stud.course_id ==
                                   db.degree_courses.course_id),
              db.majors.on(db.grad_stud.major_id == db.majors.major_id)])

    current_day = custom_strftime('{S}', dt.now())
    current_month_year = custom_strftime('%B, %Y', dt.now())

    registrar = db((db.registrar.registrar_position == 'Registrar II')).select(
        db.registrar.registrar_name)

    registrar = db(
        (db.registrar.registrar_position == 'University Registrar')).select(
        db.registrar.registrar_name)

    return locals()


def page2():
    current_day = custom_strftime('{S}', dt.now())
    current_month_year = custom_strftime('%B, %Y', dt.now())

    header = db(db.campus_college_institute.cci_id == 10).select(
        db.campus_college_institute.cci_name,
        db.campus_college_institute.address,
        db.campus_college_institute.tel_no)

    student = db(db.undergrad_stud.student_id == 1).select(
        db.undergrad_stud.first_name,
        db.undergrad_stud.middle_name,
        db.undergrad_stud.last_name,
        db.undergrad_stud.gender,
        db.undergrad_stud.date_graduated,
        db.degree_courses.course_name,
        db.degree_courses.course_abbrev,
        db.majors.major,
        left=[db.degree_courses.on(db.undergrad_stud.course_id ==
                                   db.degree_courses.course_id),
              db.majors.on(db.undergrad_stud.major_id == db.majors.major_id)])

    date_graduated = custom_strftime('%B %d, %Y', student[
        0].undergrad_stud.date_graduated)

    registrar = db((db.registrar.registrar_position == 'Registrar II')).select(
        db.registrar.registrar_name)
    return locals()


def page3():
    return locals()


def page4():
    student = db(db.undergrad_stud.student_id == 3).select(
        db.undergrad_stud.first_name,
        db.undergrad_stud.middle_name,
        db.undergrad_stud.last_name,
        db.undergrad_stud.gender,
        db.undergrad_stud.date_graduated,
        db.degree_courses.course_name,
        db.degree_courses.course_abbrev,
        db.majors.major,
        db.awards.award_title,
        left=[db.degree_courses.on(db.undergrad_stud.course_id ==
                                   db.degree_courses.course_id),
              db.majors.on(db.undergrad_stud.major_id == db.majors.major_id),
              db.awards.on(db.undergrad_stud.course_id == db.awards.award_id)])

    date_graduated = custom_strftime('%B %d, %Y', student[
        0].undergrad_stud.date_graduated)

    registrar = db(
        (db.registrar.registrar_position == 'University Registrar')).select(
        db.registrar.registrar_name)

    return locals()
    return locals()


def page5():
    current_day = custom_strftime('{S}', dt.now())
    current_month_year = custom_strftime('%B, %Y', dt.now())

    header = db(db.campus_college_institute.cci_id == 10).select(
        db.campus_college_institute.cci_name,
        db.campus_college_institute.address,
        db.campus_college_institute.tel_no)

    student = db((db.undergrad_stud.first_name == 'Francia') |
                 (db.undergrad_stud.last_name == 'Llaban')).select(
        db.undergrad_stud.first_name,
        db.undergrad_stud.middle_name,
        db.undergrad_stud.last_name,
        db.undergrad_stud.gender,
        db.undergrad_stud.date_graduated,
        db.degree_courses.course_name,
        db.degree_courses.course_abbrev,
        db.majors.major,
        left=[db.degree_courses.on(db.undergrad_stud.course_id ==
                                   db.degree_courses.course_id),
              db.majors.on(db.undergrad_stud.major_id == db.majors.major_id)])

    date_graduated = custom_strftime('%B %d, %Y', student[
        0].undergrad_stud.date_graduated)

    registrar = db((db.registrar.registrar_position == 'Registrar II')).select(
        db.registrar.registrar_name)

    return locals()


def page6():
    current_date = str(datetime.datetime.today().strftime('%B %d, %Y'))

    header = db(db.campus_college_institute.cci_id == 10).select(
        db.campus_college_institute.cci_name,
        db.campus_college_institute.address,
        db.campus_college_institute.tel_no)

    student = db((db.undergrad_stud.first_name == 'Mark Anthony') |
                 (db.undergrad_stud.last_name == 'Balla')).select(
        db.undergrad_stud.first_name,
        db.undergrad_stud.middle_name,
        db.undergrad_stud.last_name,
        db.undergrad_stud.gender,
        db.undergrad_stud.date_graduated,
        db.degree_courses.course_name,
        db.degree_courses.course_abbrev,
        db.majors.major,
        left=[db.degree_courses.on(db.undergrad_stud.course_id ==
                                   db.degree_courses.course_id),
              db.majors.on(db.undergrad_stud.major_id == db.majors.major_id)])
                 
    receipt = db(db.receipt.or_no == 1).select(
        db.receipt.or_no,
        db.receipt.date,
        db.receipt.amount,
        db.semester.sem,
        db.semester.acad_year,
        left=[db.semester.on(db.receipt.sem_id == db.semester.sem_id),
              db.undergrad_stud.on(db.receipt.student_id == db.undergrad_stud.student_id)
              ])

    or_no = f'{receipt[0].receipt.or_no:07d}'
    receipt_date = custom_strftime('%m-%d-%y', receipt[0].receipt.date)

    registrar = db(
        (db.registrar.registrar_position == 'University Registrar')).select(
        db.registrar.registrar_position,
        db.registrar.registrar_name
    )

    return locals()
