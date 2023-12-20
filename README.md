import difflib

def show_differences(list1, list2):
    differ = difflib.Differ()
    diff = list(differ.compare(list1, list2))

    print("\n".join(diff))

# Example usage:
list1 = [1, 2, 3, 4, 5]
list2 = [1, 2, 3, 6, 7]

show_differences(list1, list2)
