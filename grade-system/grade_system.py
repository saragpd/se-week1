def get_grade(mark):
    if 90 <= mark <= 100:
        return "A"
    elif 80 <= mark < 90:
        return "B"
    elif 70 <= mark < 80:
        return "C"
    elif 60 <= mark < 70:
        return "D"
    else:
        return "E"


def main():
    try:
        mark = float(input("Enter your mark (0-100): "))

        if mark < 0 or mark > 100:
            print("Invalid mark. Please enter a number between 0 and 100.")
            return

        grade = get_grade(mark)

        print(f"Entered mark: {mark}")
        print(f"Grade: {grade}")

    except ValueError:
        print("Invalid input. Please enter a number.")


if __name__ == "__main__":
    main()