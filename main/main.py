import pandas as pd
from os import mkdir
from datetime import datetime
import os


def main():
    EXCEL_NAME = "./informations.xlsx"
    EMAIL_NAME = "./furkan_email.txt"
    DIRECTION_NAME = datetime.today().strftime("%Y-%m-%d")
    # TITLE
    TITLE = "Internship Inquiry - [field]"
    #
    PROF_NAME = "[prof name]"
    MY_STUDIES = "[my studies]"
    PROF_INTERESTS = "[prof interests]"
    TITLE_NAME = "[field]"

    prof_name = ""
    my_studies = ""
    prof_interests = ""

    email = ""
    email_changed = ""

    title = "Internship Inquery - [field]"
    title_changed = ""

    starting_point = 0
    try:
        walker = list(os.walk("."))[1:]
        for tuple_walker in walker:
            count_files = len(tuple_walker[-1])
            starting_point += count_files
    except IndexError:
        pass

    with open(EMAIL_NAME, "r") as email_text:
        email = "".join(email_text.readlines())

    df = pd.read_excel(EXCEL_NAME)
    df = df.iloc[starting_point:]
    if df.shape[0] != 0:
        try:
            mkdir(DIRECTION_NAME)
        except FileExistsError:

            counter = len(list(os.walk(".")))
            DIRECTION_NAME += f" ({counter})"
            mkdir(DIRECTION_NAME)

        finally:
            DIRECTION_NAME += "/"
        for index, row in df.iterrows():
            email_name = f"{DIRECTION_NAME}{index + 1} Email.txt"
            email_changed = email
            title_changed = title

            prof_name = row["Professor Name"]
            my_studies = row["My Studies"]
            prof_interests = row["Professor Interest"]
            internship_field = row["Internship Field"]

            email_changed = email_changed.replace(PROF_NAME, prof_name)
            email_changed = email_changed.replace(MY_STUDIES, my_studies)
            email_changed = email_changed.replace(PROF_INTERESTS, prof_interests)

            title_changed = title_changed.replace(TITLE_NAME, internship_field)

            with open(email_name, "w") as email_to_save:
                email_to_save.write(f"\n{title_changed}\n\n\n")
                email_to_save.write(email_changed)

        print(
            f"""
            {index + 1 - starting_point} emails are created!
            """
        )
    else:
        print("End of your Excel, nothing is added")


if __name__ == "__main__":
    main()
