class Group:
    def __init__(self, name, num_students, subjects, subgroups):
        self.name = name
        self.num_students = num_students
        self.subjects = subjects  # List of subject names
        self.subgroups = subgroups  # List of subgroup names


class Lecturer:
    def __init__(self, name, subjects_can_teach):
        self.name = name
        self.subjects_can_teach = subjects_can_teach  # Dict of subject: [types of classes]


class Room:
    def __init__(self, name, capacity):
        self.name = name
        self.capacity = capacity


class Subject:
    def __init__(self, name, total_hours, lecture_hours, practical_hours, needs_subgroup):
        self.name = name
        self.total_hours = total_hours
        self.lecture_hours = lecture_hours
        self.practical_hours = practical_hours
        self.needs_subgroup = needs_subgroup  # Boolean


# Define the class for individual classes in the schedule
class ClassInstance:
    def __init__(self, groups, subject_name, class_type):
        self.groups = groups  # List of group or subgroup names
        self.subject_name = subject_name
        self.class_type = class_type  # 'lecture' or 'practical'
        self.lecturer = None
        self.room = None
        self.slot = None  # To be assigned during scheduling

    def __str__(self):
        return f'{self.groups} {self.subject_name} {self.class_type} {self.lecturer} {self.room} {self.slot}'
