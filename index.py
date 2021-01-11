from task.merge import Merge
from task.sum import Sum


def create_tasks():
    tasks = []

    sum = Sum("to_sum")
    tasks.append(sum)

    merge = Merge("to_merge")
    tasks.append(merge)

    return tasks


def main():
    """
    run task one by one
    """
    tasks = create_tasks()

    for task in tasks:
        task.run()


if __name__ == "__main__":
    main()
