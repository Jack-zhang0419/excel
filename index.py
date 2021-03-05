from task.merge import Merge
from task.sum import Sum
from task.combine import Combine


def create_tasks():
    tasks = []

    # tasks.append(Sum("to_sum"))
    # tasks.append(Merge("to_merge"))
    tasks.append(Combine("to_combine"))

    return tasks


def main():
    """
    run task one by one
    """
    tasks = create_tasks()

    for task in tasks:
        if task.is_valid_task():
            task.run()


if __name__ == "__main__":
    main()
