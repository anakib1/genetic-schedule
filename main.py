import copy
import random

from pretty_print import *

NUM_SLOTS = 20
WEEKS = 14
POPULATION_SIZE = 10
GENERATIONS = 10


class Schedule:
    def __init__(self):
        self.slots = [[] for _ in range(NUM_SLOTS)]
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
            taught_hours = subject_hours.get(subject_name, 0) * WEEKS
            penalty += abs(required_hours - taught_hours) ** 4

        self.fitness = penalty

    def mutate(self):

        slot1, slot2 = random.sample(range(NUM_SLOTS), 2)
        if not self.slots[slot1] or not self.slots[slot2]:
            return
        cls1 = random.choice(self.slots[slot1])
        cls2 = random.choice(self.slots[slot2])

        if self.can_swap_classes(cls1, cls2, slot1, slot2):

            self.slots[slot1].remove(cls1)
            self.slots[slot2].remove(cls2)
            self.slots[slot1].append(cls2)
            self.slots[slot2].append(cls1)
        else:

            pass

    def can_swap_classes(self, cls1, cls2, slot1, slot2):

        if not self.can_schedule_class_in_slot(cls1, slot2, exclude_classes=[cls2]):
            return False

        if not self.can_schedule_class_in_slot(cls2, slot1, exclude_classes=[cls1]):
            return False
        return True

    def can_schedule_class_in_slot(self, cls, slot_index, exclude_classes=[]):

        groups = cls['groups']
        lecturer = cls['lecturer']
        room = cls['room']

        for existing_cls in self.slots[slot_index]:
            if existing_cls in exclude_classes:
                continue

            if any(group in existing_cls['groups'] for group in groups):
                return False

            if lecturer == existing_cls['lecturer']:
                return False

            if room == existing_cls['room']:
                return False
        return True

    def local_search(self, groups, lecturers, rooms, subjects):
        """
        Perform local search to improve the schedule by making small adjustments.
        """

        for _ in range(10):

            slot_indices_with_classes = [i for i, slot in enumerate(self.slots) if slot]
            if not slot_indices_with_classes:
                break
            slot_index = random.choice(slot_indices_with_classes)
            cls = random.choice(self.slots[slot_index])

            available_slots = set(range(NUM_SLOTS)) - {slot_index}

            possible_slots = [s for s in available_slots if self.can_schedule_class_in_slot(cls, s)]
            if possible_slots:
                new_slot = random.choice(possible_slots)
                self.slots[slot_index].remove(cls)
                self.slots[new_slot].append(cls)

                cls['slot'] = new_slot
                break


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
        scheduled_class_instances = []

        for cls in class_instances:

            total_students = sum([get_group_size(group_name, groups) for group_name in cls.groups])
            possible_rooms = [room for room in rooms if room.capacity >= total_students]
            if not possible_rooms:
                continue
            cls.possible_rooms = [room.name for room in possible_rooms]
            scheduled_class_instances.append(cls)

        slots = [[] for _ in range(NUM_SLOTS)]

        group_availability = {}
        for group in groups:
            group_availability[group.name] = set(range(NUM_SLOTS))
            for subgroup_name in group.subgroups:
                group_availability[subgroup_name] = set(range(NUM_SLOTS))

        lecturer_availability = {}
        for lecturer in lecturers:
            lecturer_availability[lecturer.name] = set(range(NUM_SLOTS))

        room_availability = {}
        for room in rooms:
            room_availability[room.name] = set(range(NUM_SLOTS))

        subject_lecturers = {}
        for lecturer in lecturers:
            for subject_name, class_types in lecturer.subjects_can_teach.items():
                for class_type in class_types:
                    key = (subject_name, class_type)
                    if key not in subject_lecturers:
                        subject_lecturers[key] = []
                    subject_lecturers[key].append(lecturer.name)

        random.shuffle(scheduled_class_instances)
        for cls in scheduled_class_instances:

            key = (cls.subject_name, cls.class_type)
            possible_lecturers = subject_lecturers.get(key, [])
            if not possible_lecturers:
                continue
            random.shuffle(possible_lecturers)

            assigned = False
            for slot in random.sample(range(NUM_SLOTS), NUM_SLOTS):

                if any(slot not in group_availability[group_name] for group_name in cls.groups):
                    continue

                random.shuffle(cls.possible_rooms)
                for room_name in cls.possible_rooms:
                    if slot not in room_availability[room_name]:
                        continue

                    for lecturer_name in possible_lecturers:
                        if slot not in lecturer_availability[lecturer_name]:
                            continue

                        cls.slot = slot
                        cls.lecturer = lecturer_name
                        cls.room = room_name
                        slots[slot].append(cls)

                        for group_name in cls.groups:
                            group_availability[group_name].remove(slot)
                        lecturer_availability[lecturer_name].remove(slot)
                        room_availability[room_name].remove(slot)
                        assigned = True
                        break
                    if assigned:
                        break
                if assigned:
                    break
            if not assigned:
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
    child1.slots = copy.deepcopy(parent1.slots[:point] + parent2.slots[point:])
    child2.slots = copy.deepcopy(parent2.slots[:point] + parent1.slots[point:])
    return child1, child2


def genetic_algorithm(groups, lecturers, rooms, subjects):
    population = initial_population(groups, lecturers, rooms, subjects)
    for generation in range(GENERATIONS):

        for schedule in population:
            schedule.calculate_fitness()

        population.sort(key=lambda x: x.fitness)

        num_local_search = max(1, int(0.1 * POPULATION_SIZE))
        for i in range(num_local_search):
            population[i].local_search(groups, lecturers, rooms, subjects)

        for schedule in population[:num_local_search]:
            schedule.calculate_fitness()

        population.sort(key=lambda x: x.fitness)

        num_to_remove = max(1, int(0.2 * POPULATION_SIZE))
        population = population[:-num_to_remove]

        if generation % 5 == 0 and generation != 0:
            num_new_individuals = max(1, int(0.1 * POPULATION_SIZE))
            new_individuals = initial_population(groups, lecturers, rooms, subjects)[:num_new_individuals]
            population.extend(new_individuals)

        population = population[:POPULATION_SIZE]

        next_generation = population[:POPULATION_SIZE // 5]
        while len(next_generation) < POPULATION_SIZE:
            parent1, parent2 = random.sample(population, 2)
            child1, child2 = crossover(parent1, parent2)
            if random.random() < 0.1:
                child1.mutate()
            if random.random() < 0.1:
                child2.mutate()
            next_generation.extend([child1, child2])
        population = next_generation[:POPULATION_SIZE]

    return population[0]


if __name__ == "__main__":
    groups = load_groups('groups.csv')
    lecturers = load_lecturers('lecturers.csv')
    rooms = load_rooms('rooms.csv')
    subjects = load_subjects('subjects.csv')

    best_schedule = genetic_algorithm(groups, lecturers, rooms, subjects)

    export_schedule_to_excel(best_schedule, groups, filename='schedule.xlsx')
    print_subject_hours_report(best_schedule, groups, subjects)
