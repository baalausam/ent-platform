from django.contrib import admin
from .models import (School, SchoolClass, UserProfile, Subject,
                     Question, TestAttempt, Answer, QuestionGroup)


@admin.register(Subject)
class SubjectAdmin(admin.ModelAdmin):
    list_display = ('name', 'code', 'category', 'is_active')
    list_filter = ('category', 'is_active')


@admin.register(School)
class SchoolAdmin(admin.ModelAdmin):
    list_display = ('name', 'city', 'teacher_code', 'zavuch_code', 'student_code')
    readonly_fields = ('teacher_code', 'zavuch_code', 'student_code')


@admin.register(SchoolClass)
class SchoolClassAdmin(admin.ModelAdmin):
    list_display = ('school', 'grade', 'letter')
    list_filter = ('school', 'grade')


@admin.register(UserProfile)
class UserProfileAdmin(admin.ModelAdmin):
    list_display = ('user', 'role', 'school', 'is_subject_teacher', 'is_homeroom_teacher')
    list_filter = ('role', 'school', 'is_subject_teacher', 'is_homeroom_teacher')


@admin.register(Question)
class QuestionAdmin(admin.ModelAdmin):
    list_display = ('block', 'subject', 'question_type', 'text_short', 'status', 'author', 'times_used')
    list_filter = ('block', 'subject', 'status', 'question_type', 'author')
    search_fields = ('text',)
    readonly_fields = ('times_used', 'approved_at')

    def text_short(self, obj):
        return obj.text[:60]
    text_short.short_description = 'Текст'


@admin.register(TestAttempt)
class TestAttemptAdmin(admin.ModelAdmin):
    list_display = ('student', 'status', 'total_score', 'tab_switch_count', 'started_at')
    list_filter = ('status',)
    readonly_fields = ('started_at', 'finished_at')


@admin.register(Answer)
class AnswerAdmin(admin.ModelAdmin):
    list_display = ('attempt', 'subject_snapshot', 'chosen', 'score')
    list_filter = ('subject_snapshot',)
    from .models import QuestionGroup

@admin.register(QuestionGroup)
class QuestionGroupAdmin(admin.ModelAdmin):
    list_display = ('title', 'block', 'subject', 'author', 'status', 'created_at')
    list_filter = ('block', 'status', 'author')
    search_fields = ('title', 'context_text')