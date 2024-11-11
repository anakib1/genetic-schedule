import random
from collections import defaultdict
import copy
from pretty_print import *

NUM_SLOTS = 20
POPULATION_SIZE = 100
GENERATIONS = 500


class Schedule:
    def __init__(self):
        self.slots = [[] for _ in range(NUM_SLOTS)]  # List of lists of classes
        self.fitness = None

    def calculate_fitness(self):
        penalty = 0
        group_windows = defaultdict(list)
        lecturer_windows = defaultdict(list)
        subject_hours = defaultdict(float)

        for slot_index, slot in enumerate(self.slots):
            for cls in slot:
                subject_hours[cls['subject']] += 1.5

                for group in cls['groups']:
                    group_windows[group].append(slot_index)

                lecturer_windows[cls['lecturer']].append(slot_index)

        for times in group_windows.values():
            times.sort()
            penalty += sum((times[i + 1] - times[i] - 1) for i in range(len(times) - 1))

        for times in lecturer_windows.values():
            times.sort()
            penalty += sum((times[i + 1] - times[i] - 1) for i in range(len(times) - 1))

        for subject_name, subject in subjects.items():
            required_hours = subject.total_hours
            taught_hours = subject_hours.get(subject_name, 0)
            penalty += abs(required_hours - taught_hours)

        self.fitness = penalty

    def mutate(self):
        # Randomly swap two classes
        slot1, slot2 = random.sample(range(NUM_SLOTS), 2)
        if self.slots[slot1] and self.slots[slot2]:
            cls1 = random.choice(self.slots[slot1])
            cls2 = random.choice(self.slots[slot2])
            self.slots[slot1].remove(cls1)
            self.slots[slot2].remove(cls2)
            self.slots[slot1].append(cls2)
            self.slots[slot2].append(cls1)


def generate_class_instances(groups, subjects):
    class_instances = []
    for group in groups:
        for subject_name in group.subjects:
            subject = subjects[subject_name]
            num_lecture_sessions = int(subject.lecture_hours / 1.5)
            for _ in range(num_lecture_sessions):
                cls = ClassInstance([group.name], subject_name, 'lecture')
                class_instances.append(cls)
            num_practical_sessions = int(subject.practical_hours / 1.5)
            if subject.needs_subgroup:
                for subgroup_name in group.subgroups:
                    for _ in range(num_practical_sessions):
                        cls = ClassInstance([subgroup_name], subject_name, 'practical')
                        class_instances.append(cls)
            else:
                for _ in range(num_practical_sessions):
                    cls = ClassInstance([group.name], subject_name, 'practical')
                    class_instances.append(cls)
    return class_instances


def get_group_size(group_name, groups):
    for group in groups:
        if group.name == group_name:
            return group.num_students
        elif group_name in group.subgroups:
            return group.num_students // 2
    return 0


def initial_population(groups, lecturers, rooms, subjects):
    original_class_instances = generate_class_instances(groups, subjects)

    population = []
    for _ in range(POPULATION_SIZE):
        class_instances = copy.deepcopy(original_class_instances)
        schedule = Schedule()

        for cls in class_instances:
            possible_lecturers = [
                lec for lec in lecturers
                if cls.subject_name in lec.subjects_can_teach and cls.class_type in lec.subjects_can_teach[
                    cls.subject_name]
            ]
            if not possible_lecturers:
                continue
            cls.lecturer = random.choice(possible_lecturers).name

            total_students = sum([get_group_size(group_name, groups) for group_name in cls.groups])
            possible_rooms = [room for room in rooms if room.capacity >= total_students]
            if not possible_rooms:
                continue
            cls.room = random.choice(possible_rooms).name

        slots = [[] for _ in range(NUM_SLOTS)]
        group_availability = defaultdict(lambda: set(range(NUM_SLOTS)))
        lecturer_availability = defaultdict(lambda: set(range(NUM_SLOTS)))
        room_availability = defaultdict(lambda: set(range(NUM_SLOTS)))

        random.shuffle(class_instances)
        for cls in class_instances:
            available_slots = set(range(NUM_SLOTS))
            for group_name in cls.groups:
                available_slots &= group_availability[group_name]
            available_slots &= lecturer_availability[cls.lecturer]
            available_slots &= room_availability[cls.room]
            if available_slots:
                slot = random.choice(list(available_slots))
                cls.slot = slot
                slots[slot].append(cls)
                for group_name in cls.groups:
                    group_availability[group_name].remove(slot)
                lecturer_availability[cls.lecturer].remove(slot)
                room_availability[cls.room].remove(slot)
            else:
                continue

        for slot_index, slot_classes in enumerate(slots):
            schedule.slots[slot_index] = [
                {
                    'groups': cls.groups,
                    'subject': cls.subject_name,
                    'lecturer': cls.lecturer,
                    'room': cls.room,
                    'type': cls.class_type
                } for cls in slot_classes
            ]
        population.append(schedule)
    return population


def crossover(parent1, parent2):
    point = random.randint(1, NUM_SLOTS - 1)
    child1 = Schedule()
    child2 = Schedule()
    child1.slots = parent1.slots[:point] + parent2.slots[point:]
    child2.slots = parent2.slots[:point] + parent1.slots[point:]
    return child1, child2


def genetic_algorithm(groups, lecturers, rooms, subjects):
    population = initial_population(groups, lecturers, rooms, subjects)
    for generation in range(GENERATIONS):
        for schedule in population:
            schedule.calculate_fitness()
        population.sort(key=lambda x: x.fitness)
        next_generation = population[:POPULATION_SIZE // 5]
        while len(next_generation) < POPULATION_SIZE:
            parent1, parent2 = random.sample(population[:POPULATION_SIZE // 2], 2)
            child1, child2 = crossover(parent1, parent2)
            if random.random() < 0.1:
                child1.mutate()
            if random.random() < 0.1:
                child2.mutate()
            next_generation.extend([child1, child2])
        population = next_generation[:POPULATION_SIZE]
    return population[0]


if __name__ == "__main__":
    # Load data (you need to have appropriate CSV files or generate data)
    groups = load_groups('groups.csv')
    lecturers = load_lecturers('lecturers.csv')
    rooms = load_rooms('rooms.csv')
    subjects = load_subjects('subjects.csv')

    # Run the genetic algorithm to get the best schedule
    best_schedule = genetic_algorithm(groups, lecturers, rooms, subjects)

    export_schedule_to_excel(best_schedule, groups, filename='schedule.xlsx')
