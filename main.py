import copy
import random

from pretty_print import *

NUM_SLOTS = 20
WEEKS = 14
POPULATION_SIZE = 10
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
            taught_hours = subject_hours.get(subject_name, 0) * WEEKS
            penalty += abs(required_hours - taught_hours) ** 4

        self.fitness = penalty

    def mutate(self):
        # Randomly select two different slots
        slot1, slot2 = random.sample(range(NUM_SLOTS), 2)
        if not self.slots[slot1] or not self.slots[slot2]:
            return  # One of the slots is empty, can't swap
        cls1 = random.choice(self.slots[slot1])
        cls2 = random.choice(self.slots[slot2])

        # Check if swapping cls1 and cls2 maintains hard constraints
        if self.can_swap_classes(cls1, cls2, slot1, slot2):
            # Perform the swap
            self.slots[slot1].remove(cls1)
            self.slots[slot2].remove(cls2)
            self.slots[slot1].append(cls2)
            self.slots[slot2].append(cls1)
        else:
            # Mutation cannot be performed without breaking hard constraints
            # Optionally, try a different mutation or do nothing
            pass

    def can_swap_classes(self, cls1, cls2, slot1, slot2):
        # Check if cls1 can be scheduled in slot2
        if not self.can_schedule_class_in_slot(cls1, slot2, exclude_classes=[cls2]):
            return False
        # Check if cls2 can be scheduled in slot1
        if not self.can_schedule_class_in_slot(cls2, slot1, exclude_classes=[cls1]):
            return False
        return True

    def can_schedule_class_in_slot(self, cls, slot_index, exclude_classes=[]):
        # Check if cls can be scheduled in slot_index without conflicts
        # Exclude classes that are being moved to avoid false positives
        groups = cls['groups']
        lecturer = cls['lecturer']
        room = cls['room']

        for existing_cls in self.slots[slot_index]:
            if existing_cls in exclude_classes:
                continue
            # Check for group conflict
            if any(group in existing_cls['groups'] for group in groups):
                return False
            # Check for lecturer conflict
            if lecturer == existing_cls['lecturer']:
                return False
            # Check for room conflict
            if room == existing_cls['room']:
                return False
        return True


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
    # First, generate all class instances that need to be scheduled
    original_class_instances = generate_class_instances(groups, subjects)

    population = []
    for _ in range(POPULATION_SIZE):
        # Deep copy of class instances for this individual
        class_instances = copy.deepcopy(original_class_instances)
        schedule = Schedule()
        scheduled_class_instances = []

        # Assign possible rooms to class instances (lecturers will be assigned during slot assignment)
        for cls in class_instances:
            # Find rooms that can accommodate the group(s)
            total_students = sum([get_group_size(group_name, groups) for group_name in cls.groups])
            possible_rooms = [room for room in rooms if room.capacity >= total_students]
            if not possible_rooms:
                continue  # Cannot schedule this class
            cls.possible_rooms = [room.name for room in possible_rooms]
            scheduled_class_instances.append(cls)

        # Now, assign class instances to slots and assign lecturers considering their availability
        # Initialize availability dictionaries
        slots = [[] for _ in range(NUM_SLOTS)]

        # Initialize availability for groups and subgroups
        group_availability = {}
        for group in groups:
            group_availability[group.name] = set(range(NUM_SLOTS))
            for subgroup_name in group.subgroups:
                group_availability[subgroup_name] = set(range(NUM_SLOTS))

        # Initialize availability for lecturers
        lecturer_availability = {}
        for lecturer in lecturers:
            lecturer_availability[lecturer.name] = set(range(NUM_SLOTS))

        # Initialize availability for rooms
        room_availability = {}
        for room in rooms:
            room_availability[room.name] = set(range(NUM_SLOTS))

        # Build a mapping of subject and class type to lecturers who can teach it
        subject_lecturers = {}
        for lecturer in lecturers:
            for subject_name, class_types in lecturer.subjects_can_teach.items():
                for class_type in class_types:
                    key = (subject_name, class_type)
                    if key not in subject_lecturers:
                        subject_lecturers[key] = []
                    subject_lecturers[key].append(lecturer.name)

        # Shuffle class instances to randomize scheduling
        random.shuffle(scheduled_class_instances)
        for cls in scheduled_class_instances:
            # Get the list of possible lecturers for this class
            key = (cls.subject_name, cls.class_type)
            possible_lecturers = subject_lecturers.get(key, [])
            if not possible_lecturers:
                continue  # No lecturers available for this class
            random.shuffle(possible_lecturers)

            # Try to assign the class to an available slot, room, and lecturer
            assigned = False
            for slot in random.sample(range(NUM_SLOTS), NUM_SLOTS):
                # Check if all groups are available at this slot
                if any(slot not in group_availability[group_name] for group_name in cls.groups):
                    continue

                # Try to find an available room and lecturer for this slot
                random.shuffle(cls.possible_rooms)
                for room_name in cls.possible_rooms:
                    if slot not in room_availability[room_name]:
                        continue
                    # Try to find an available lecturer for this slot
                    for lecturer_name in possible_lecturers:
                        if slot not in lecturer_availability[lecturer_name]:
                            continue
                        # Assign lecturer, room, and slot
                        cls.slot = slot
                        cls.lecturer = lecturer_name
                        cls.room = room_name
                        slots[slot].append(cls)
                        # Update availabilities
                        for group_name in cls.groups:
                            group_availability[group_name].remove(slot)
                        lecturer_availability[lecturer_name].remove(slot)
                        room_availability[room_name].remove(slot)
                        assigned = True
                        break  # Lecturer assigned, move to next class
                    if assigned:
                        break  # Room assigned, move to next class
                if assigned:
                    break  # Class assigned, move to next class
            if not assigned:
                continue  # Could not assign this class due to availability constraints

        # Assign slots to schedule
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
    print_subject_hours_report(best_schedule, groups, subjects)
