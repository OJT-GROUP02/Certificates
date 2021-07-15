--
-- PostgreSQL database dump
--

-- Dumped from database version 13.3
-- Dumped by pg_dump version 13.3

SET statement_timeout = 0;
SET lock_timeout = 0;
SET idle_in_transaction_session_timeout = 0;
SET client_encoding = 'UTF8';
SET standard_conforming_strings = on;
SELECT pg_catalog.set_config('search_path', '', false);
SET check_function_bodies = false;
SET xmloption = content;
SET client_min_messages = warning;
SET row_security = off;

SET default_tablespace = '';

SET default_table_access_method = heap;

--
-- Name: awards; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.awards (
    award_id integer NOT NULL,
    award_title character varying(50) NOT NULL
);


ALTER TABLE public.awards OWNER TO postgres;

--
-- Name: awards_award_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

CREATE SEQUENCE public.awards_award_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER TABLE public.awards_award_id_seq OWNER TO postgres;

--
-- Name: awards_award_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: postgres
--

ALTER SEQUENCE public.awards_award_id_seq OWNED BY public.awards.award_id;


--
-- Name: campus_college_institute; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.campus_college_institute (
    cci_id integer NOT NULL,
    cci_name character varying(200) NOT NULL,
    campus character varying(100),
    address character varying(200) NOT NULL,
    tel_no character varying(30)
);


ALTER TABLE public.campus_college_institute OWNER TO postgres;

--
-- Name: campus_college_institute_cci_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

CREATE SEQUENCE public.campus_college_institute_cci_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER TABLE public.campus_college_institute_cci_id_seq OWNER TO postgres;

--
-- Name: campus_college_institute_cci_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: postgres
--

ALTER SEQUENCE public.campus_college_institute_cci_id_seq OWNED BY public.campus_college_institute.cci_id;


--
-- Name: degree_courses; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.degree_courses (
    course_id bigint NOT NULL,
    course_name character varying(200) NOT NULL,
    course_abbrev character varying(20),
    cci_id integer NOT NULL
);


ALTER TABLE public.degree_courses OWNER TO postgres;

--
-- Name: degree_courses_course_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

CREATE SEQUENCE public.degree_courses_course_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER TABLE public.degree_courses_course_id_seq OWNER TO postgres;

--
-- Name: degree_courses_course_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: postgres
--

ALTER SEQUENCE public.degree_courses_course_id_seq OWNED BY public.degree_courses.course_id;


--
-- Name: grad_stud; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.grad_stud (
    student_id bigint NOT NULL,
    first_name character varying(50) NOT NULL,
    middle_name character varying(50),
    last_name character varying(50) NOT NULL,
    gender character(1) NOT NULL,
    course_id integer NOT NULL,
    major_id integer
);


ALTER TABLE public.grad_stud OWNER TO postgres;

--
-- Name: grad_stud_student_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

CREATE SEQUENCE public.grad_stud_student_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER TABLE public.grad_stud_student_id_seq OWNER TO postgres;

--
-- Name: grad_stud_student_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: postgres
--

ALTER SEQUENCE public.grad_stud_student_id_seq OWNED BY public.grad_stud.student_id;


--
-- Name: majors; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.majors (
    major_id bigint NOT NULL,
    course_id integer NOT NULL,
    major character varying(200) NOT NULL
);


ALTER TABLE public.majors OWNER TO postgres;

--
-- Name: majors_major_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

CREATE SEQUENCE public.majors_major_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER TABLE public.majors_major_id_seq OWNER TO postgres;

--
-- Name: majors_major_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: postgres
--

ALTER SEQUENCE public.majors_major_id_seq OWNED BY public.majors.major_id;


--
-- Name: receipt; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.receipt (
    or_no integer NOT NULL,
    student_id integer NOT NULL,
    sem_id integer NOT NULL,
    amount numeric(7,2),
    date date
);


ALTER TABLE public.receipt OWNER TO postgres;

--
-- Name: receipt_or_no_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

CREATE SEQUENCE public.receipt_or_no_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER TABLE public.receipt_or_no_seq OWNER TO postgres;

--
-- Name: receipt_or_no_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: postgres
--

ALTER SEQUENCE public.receipt_or_no_seq OWNED BY public.receipt.or_no;


--
-- Name: registrar; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.registrar (
    registrar_id integer NOT NULL,
    registrar_name character varying(150) NOT NULL,
    registrar_position character varying(50) NOT NULL
);


ALTER TABLE public.registrar OWNER TO postgres;

--
-- Name: registrar_registrar_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

CREATE SEQUENCE public.registrar_registrar_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER TABLE public.registrar_registrar_id_seq OWNER TO postgres;

--
-- Name: registrar_registrar_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: postgres
--

ALTER SEQUENCE public.registrar_registrar_id_seq OWNED BY public.registrar.registrar_id;


--
-- Name: semester; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.semester (
    sem_id bigint NOT NULL,
    sem character varying(20) NOT NULL,
    acad_year character varying(20) NOT NULL
);


ALTER TABLE public.semester OWNER TO postgres;

--
-- Name: semester_sem_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

CREATE SEQUENCE public.semester_sem_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER TABLE public.semester_sem_id_seq OWNER TO postgres;

--
-- Name: semester_sem_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: postgres
--

ALTER SEQUENCE public.semester_sem_id_seq OWNED BY public.semester.sem_id;


--
-- Name: undergrad_stud; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.undergrad_stud (
    student_id bigint NOT NULL,
    first_name character varying(50) NOT NULL,
    middle_name character varying(50),
    last_name character varying(50) NOT NULL,
    gender character(1) NOT NULL,
    gwa numeric(3,2),
    award_id integer,
    date_graduated date,
    course_id integer NOT NULL,
    major_id integer
);


ALTER TABLE public.undergrad_stud OWNER TO postgres;

--
-- Name: undergrad_stud_student_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

CREATE SEQUENCE public.undergrad_stud_student_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER TABLE public.undergrad_stud_student_id_seq OWNER TO postgres;

--
-- Name: undergrad_stud_student_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: postgres
--

ALTER SEQUENCE public.undergrad_stud_student_id_seq OWNED BY public.undergrad_stud.student_id;


--
-- Name: awards award_id; Type: DEFAULT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.awards ALTER COLUMN award_id SET DEFAULT nextval('public.awards_award_id_seq'::regclass);


--
-- Name: campus_college_institute cci_id; Type: DEFAULT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.campus_college_institute ALTER COLUMN cci_id SET DEFAULT nextval('public.campus_college_institute_cci_id_seq'::regclass);


--
-- Name: degree_courses course_id; Type: DEFAULT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.degree_courses ALTER COLUMN course_id SET DEFAULT nextval('public.degree_courses_course_id_seq'::regclass);


--
-- Name: grad_stud student_id; Type: DEFAULT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.grad_stud ALTER COLUMN student_id SET DEFAULT nextval('public.grad_stud_student_id_seq'::regclass);


--
-- Name: majors major_id; Type: DEFAULT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.majors ALTER COLUMN major_id SET DEFAULT nextval('public.majors_major_id_seq'::regclass);


--
-- Name: receipt or_no; Type: DEFAULT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.receipt ALTER COLUMN or_no SET DEFAULT nextval('public.receipt_or_no_seq'::regclass);


--
-- Name: registrar registrar_id; Type: DEFAULT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.registrar ALTER COLUMN registrar_id SET DEFAULT nextval('public.registrar_registrar_id_seq'::regclass);


--
-- Name: semester sem_id; Type: DEFAULT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.semester ALTER COLUMN sem_id SET DEFAULT nextval('public.semester_sem_id_seq'::regclass);


--
-- Name: undergrad_stud student_id; Type: DEFAULT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.undergrad_stud ALTER COLUMN student_id SET DEFAULT nextval('public.undergrad_stud_student_id_seq'::regclass);


--
-- Data for Name: awards; Type: TABLE DATA; Schema: public; Owner: postgres
--

COPY public.awards (award_id, award_title) FROM stdin;
1	Summa Cum Laude
2	Magna Cum Laude
3	Cum Laude
4	with Academic Distinction
\.


--
-- Data for Name: campus_college_institute; Type: TABLE DATA; Schema: public; Owner: postgres
--

COPY public.campus_college_institute (cci_id, cci_name, campus, address, tel_no) FROM stdin;
1	College of Arts and Letters	Main Campus	Legazpi City, Albay	\N
2	College of Education	Main Campus	Daraga, Albay	\N
3	College of Nursing	Main Campus	Legazpi City, Albay	\N
4	College of Science	Main Campus	Legazpi City, Albay	\N
5	Institute of Physical Education, Sports, and Recreation	Main Campus	Daraga, Albay	\N
6	College of Business, Economics, and Management	Daraga Campus	Daraga, Albay	\N
7	College of Social Sciences and Philosophy	Daraga Campus	Daraga, Albay	\N
8	College of Engineering	East Campus	Legazpi City, Albay	\N
9	Institute of Architecture	East Campus	Legazpi City, Albay	\N
10	College of Industrial Technology	East Campus	Legazpi City, Albay	(052) 820-8277
11	Graduate School		Legazpi City, Albay	(052) 481-7881
12	College of Agriculture and Forestry	Guinobatan Campus	Guinobatan, Albay	\N
13	Tabaco Campus	Tabaco Campus	Tabaco, Albay	\N
14	Polangui Campus	Polangui Campus	Polangui, Albay	\N
15	Gubat Campus	Gubat Campus	Gubat, Sorsogon	\N
\.


--
-- Data for Name: degree_courses; Type: TABLE DATA; Schema: public; Owner: postgres
--

COPY public.degree_courses (course_id, course_name, course_abbrev, cci_id) FROM stdin;
1	Master of Arts in Education	MAED	11
2	Bachelor of Science in Nursing	BSN	3
3	Bachelor of Science in Industrial Education	BSIE	10
4	Bachelor of Science in Mechanical Technology	BSMT	10
\.


--
-- Data for Name: grad_stud; Type: TABLE DATA; Schema: public; Owner: postgres
--

COPY public.grad_stud (student_id, first_name, middle_name, last_name, gender, course_id, major_id) FROM stdin;
1	Noemi	Bane	Lumbao	F	1	1
\.


--
-- Data for Name: majors; Type: TABLE DATA; Schema: public; Owner: postgres
--

COPY public.majors (major_id, course_id, major) FROM stdin;
1	1	Administration and Supervision
2	3	Food and Service Management
3	3	Food Technology
\.


--
-- Data for Name: receipt; Type: TABLE DATA; Schema: public; Owner: postgres
--

COPY public.receipt (or_no, student_id, sem_id, amount, date) FROM stdin;
1	4	1	75.00	2010-06-16
\.


--
-- Data for Name: registrar; Type: TABLE DATA; Schema: public; Owner: postgres
--

COPY public.registrar (registrar_id, registrar_name, registrar_position) FROM stdin;
1	Corazon N. Bazar	University Registrar
2	Sophia A. Romero	Registrar II
3	Sophia A. Romero	University Registrar
4	Catherine D. Lagata	Registrar II
\.


--
-- Data for Name: semester; Type: TABLE DATA; Schema: public; Owner: postgres
--

COPY public.semester (sem_id, sem, acad_year) FROM stdin;
1	2nd Semester	2009-2010
\.


--
-- Data for Name: undergrad_stud; Type: TABLE DATA; Schema: public; Owner: postgres
--

COPY public.undergrad_stud (student_id, first_name, middle_name, last_name, gender, gwa, award_id, date_graduated, course_id, major_id) FROM stdin;
1	Irvin	Montalba	Sandia	M	\N	\N	2002-03-24	2	\N
2	Imee Janine	Ordonia	Abalon	F	2.01	\N	\N	3	2
3	Francia	Quinones	Llaban	F	\N	3	1997-04-03	3	3
4	Mark Anthony	Paras	Balla	M	\N	\N	\N	4	\N
\.


--
-- Name: awards_award_id_seq; Type: SEQUENCE SET; Schema: public; Owner: postgres
--

SELECT pg_catalog.setval('public.awards_award_id_seq', 4, true);


--
-- Name: campus_college_institute_cci_id_seq; Type: SEQUENCE SET; Schema: public; Owner: postgres
--

SELECT pg_catalog.setval('public.campus_college_institute_cci_id_seq', 15, true);


--
-- Name: degree_courses_course_id_seq; Type: SEQUENCE SET; Schema: public; Owner: postgres
--

SELECT pg_catalog.setval('public.degree_courses_course_id_seq', 4, true);


--
-- Name: grad_stud_student_id_seq; Type: SEQUENCE SET; Schema: public; Owner: postgres
--

SELECT pg_catalog.setval('public.grad_stud_student_id_seq', 1, true);


--
-- Name: majors_major_id_seq; Type: SEQUENCE SET; Schema: public; Owner: postgres
--

SELECT pg_catalog.setval('public.majors_major_id_seq', 3, true);


--
-- Name: receipt_or_no_seq; Type: SEQUENCE SET; Schema: public; Owner: postgres
--

SELECT pg_catalog.setval('public.receipt_or_no_seq', 1, true);


--
-- Name: registrar_registrar_id_seq; Type: SEQUENCE SET; Schema: public; Owner: postgres
--

SELECT pg_catalog.setval('public.registrar_registrar_id_seq', 4, true);


--
-- Name: semester_sem_id_seq; Type: SEQUENCE SET; Schema: public; Owner: postgres
--

SELECT pg_catalog.setval('public.semester_sem_id_seq', 1, true);


--
-- Name: undergrad_stud_student_id_seq; Type: SEQUENCE SET; Schema: public; Owner: postgres
--

SELECT pg_catalog.setval('public.undergrad_stud_student_id_seq', 4, true);


--
-- Name: awards awards_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.awards
    ADD CONSTRAINT awards_pkey PRIMARY KEY (award_id);


--
-- Name: campus_college_institute campus_college_institute_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.campus_college_institute
    ADD CONSTRAINT campus_college_institute_pkey PRIMARY KEY (cci_id);


--
-- Name: degree_courses degree_courses_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.degree_courses
    ADD CONSTRAINT degree_courses_pkey PRIMARY KEY (course_id);


--
-- Name: grad_stud grad_stud_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.grad_stud
    ADD CONSTRAINT grad_stud_pkey PRIMARY KEY (student_id);


--
-- Name: majors majors_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.majors
    ADD CONSTRAINT majors_pkey PRIMARY KEY (major_id);


--
-- Name: receipt receipt_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.receipt
    ADD CONSTRAINT receipt_pkey PRIMARY KEY (or_no);


--
-- Name: semester semester_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.semester
    ADD CONSTRAINT semester_pkey PRIMARY KEY (sem_id);


--
-- Name: undergrad_stud undergrad_stud_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.undergrad_stud
    ADD CONSTRAINT undergrad_stud_pkey PRIMARY KEY (student_id);


--
-- Name: degree_courses degree_courses_cci_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.degree_courses
    ADD CONSTRAINT degree_courses_cci_id_fkey FOREIGN KEY (cci_id) REFERENCES public.campus_college_institute(cci_id);


--
-- Name: grad_stud grad_stud_course_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.grad_stud
    ADD CONSTRAINT grad_stud_course_id_fkey FOREIGN KEY (course_id) REFERENCES public.degree_courses(course_id);


--
-- Name: grad_stud grad_stud_major_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.grad_stud
    ADD CONSTRAINT grad_stud_major_id_fkey FOREIGN KEY (major_id) REFERENCES public.majors(major_id);


--
-- Name: majors majors_course_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.majors
    ADD CONSTRAINT majors_course_id_fkey FOREIGN KEY (course_id) REFERENCES public.degree_courses(course_id);


--
-- Name: receipt receipt_sem_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.receipt
    ADD CONSTRAINT receipt_sem_id_fkey FOREIGN KEY (sem_id) REFERENCES public.semester(sem_id);


--
-- Name: receipt receipt_student_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.receipt
    ADD CONSTRAINT receipt_student_id_fkey FOREIGN KEY (student_id) REFERENCES public.undergrad_stud(student_id);


--
-- Name: undergrad_stud undergrad_stud_course_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.undergrad_stud
    ADD CONSTRAINT undergrad_stud_course_id_fkey FOREIGN KEY (course_id) REFERENCES public.degree_courses(course_id);


--
-- Name: undergrad_stud undergrad_stud_major_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.undergrad_stud
    ADD CONSTRAINT undergrad_stud_major_id_fkey FOREIGN KEY (major_id) REFERENCES public.majors(major_id);


--
-- PostgreSQL database dump complete
--

