from task.merge import Merge
from task.sum import Sum
from task.merge_ab import MergeAB


def create_tasks():
    tasks = []

    # tasks.append(Sum("to_sum"))
    # tasks.append(Merge("to_merge"))
    tasks.append(MergeAB("to_merge_AB"))

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
