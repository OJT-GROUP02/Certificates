# ivee
db = DAL('postgres://postgres:april17@localhost/newdb')

#maedel
# db = DAL('postgres://postgres:1612@localhost/certification_db')

#vega
# db = DAL('postgres://postgres:postgres@localhost/certificate')

db.define_table('awards',
                Field('award_id'),
                Field('award_title'),
                migrate=False
                )

db.define_table('campus_college_institute',
                Field('cci_id'),
                Field('cci_name'),
                Field('campus'),
                Field('address'),
                Field('tel_no'),
                migrate=False
                )

db.define_table('semester',
                Field('sem_id'),
                Field('sem'),
                Field('acad_year'),
                migrate=False
                )

db.define_table('registrar',
                Field('registrar_id'),
                Field('registrar_name'),
                Field('registrar_position'),
                migrate=False
                )

db.define_table('degree_courses',
                Field('course_id'),
                Field('course_name'),
                Field('course_abbrev'),
                Field('cci_id'),
                migrate=False
                )

db.define_table('majors',
                Field('major_id'),
                Field('course_id'),
                Field('major'),
                migrate=False
                )

db.define_table('undergrad_stud',
                Field('student_id'),
                Field('first_name'),
                Field('middle_name'),
                Field('last_name'),
                Field('gender'),
                Field('gwa'),
                Field('award_id'),
                Field('date_graduated'),
                Field('course_id'),
                Field('major_id'),
                migrate=False
                )

db.define_table('grad_stud',
                Field('student_id'),
                Field('first_name'),
                Field('middle_name'),
                Field('last_name'),
                Field('gender'),
                Field('course_id'),
                Field('major_id'),
                migrate=False
                )

db.define_table('receipt',
                Field('or_no'),
                Field('student_id'),
                Field('sem_id'),
                Field('amount'),
                Field('date'),
                migrate=False
                )