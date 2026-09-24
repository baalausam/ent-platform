from django.urls import path
from . import views

app_name = 'core'

urlpatterns = [
    path('', views.home, name='home'),
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),

    path('register/director/', views.register_director, name='register_director'),
    path('register/teacher/', views.register_teacher, name='register_teacher'),
    path('register/student/', views.register_student, name='register_student'),

    path('director/setup-classes/', views.director_setup_classes, name='director_setup_classes'),
    path('director/dashboard/', views.director_dashboard, name='director_dashboard'),
    path('director/questions/', views.review_questions, name='review_questions'),
    path('director/questions/<int:question_id>/approve/', views.approve_question, name='approve_question'),
    path('director/questions/<int:question_id>/reject/', views.reject_question, name='reject_question'),
    path('director/groups/<int:group_id>/approve/', views.approve_group, name='approve_group'),
    path('director/groups/<int:group_id>/reject/', views.reject_group, name='reject_group'),
    path('director/results/', views.school_results, name='school_results'),
    path('director/results/export/', views.export_results_excel, name='export_results_excel'),
    path('director/results/<int:session_id>/', views.school_results_detail, name='school_results_detail'),

    path('zavuch/dashboard/', views.zavuch_dashboard, name='zavuch_dashboard'),

    path('teacher/dashboard/', views.teacher_dashboard, name='teacher_dashboard'),
    path('teacher/questions/', views.my_questions, name='my_questions'),
    path('teacher/questions/add/', views.add_question, name='add_question'),
    path('teacher/questions/<int:question_id>/submit/', views.submit_question_for_review, name='submit_question_for_review'),
    path('teacher/questions/<int:question_id>/cancel/', views.cancel_question, name='cancel_question'),
    path('teacher/groups/', views.my_question_groups, name='my_question_groups'),
    path('teacher/groups/add/', views.add_question_group, name='add_question_group'),
    path('teacher/groups/<int:group_id>/submit/', views.submit_group_for_review, name='submit_group_for_review'),
    path('teacher/groups/<int:group_id>/cancel/', views.cancel_group, name='cancel_group'),
    path('teacher/class-results/', views.class_results, name='class_results'),
    path('teacher/class-results/<int:session_id>/', views.class_results_detail, name='class_results_detail'),

    path('sessions/', views.manage_sessions, name='manage_sessions'),
    path('sessions/<int:session_id>/toggle/', views.toggle_session, name='toggle_session'),
    path('sessions/<int:session_id>/delete/', views.delete_session, name='delete_session'),

    path('student/dashboard/', views.student_dashboard, name='student_dashboard'),
    path('student/start-test/', views.start_test, name='start_test'),
    path('student/test/<int:attempt_id>/', views.take_test, name='take_test'),
    path('student/test/<int:attempt_id>/finish/', views.finish_test, name='finish_test'),
    path('student/test/<int:attempt_id>/result/', views.test_result, name='test_result'),

    path('api/test/<int:attempt_id>/answer/<int:answer_id>/', views.save_answer, name='save_answer'),
    path('api/test/<int:attempt_id>/violation/', views.register_violation, name='register_violation'),

    path('school/students/', views.students_list, name='students_list'),
    path('school/teachers/', views.teachers_list, name='teachers_list'),
    path('school/class/<int:class_id>/students/', views.class_students_list, name='class_students_list'),
    path('profile/edit/', views.edit_profile, name='edit_profile'),
    path('question/<int:question_id>/', views.question_detail, name='question_detail'),
    path('group/<int:group_id>/', views.group_detail, name='group_detail'),
]