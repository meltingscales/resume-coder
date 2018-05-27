import os
import yaml

from docx import Document


def islist(thing):
    return isinstance(thing, list)


def isdict(thing):
    return isinstance(thing, dict)


def isstr(thing):
    return isinstance(thing, str)


def yml_to_dict(path: str) -> dict:
    with open(path, 'r') as f:
        return yaml.safe_load(f)


class ResumeCoder:
    document = Document()
    path = ""
    data = {}

    def __init__(self, path):
        self.path = path

    def write(self, path="resume.docx"):
        """Save document to file."""
        self.document.save(os.path.join(self.path, path))


def contact_info(resumecoder: ResumeCoder, path: str, sep=" | "):
    def name(name, doc):
        doc.add_paragraph(name)

    def email(email, doc):

        p = doc.add_paragraph('')

        if isstr(email):  # just one email
            p.add_run(email)

        elif islist(email):  # list of emails
            for i in range(len(email)):
                p.add_run(email[i])

                if i < len(email) - 1:  # prevent sep at end of list
                    p.add_run(sep)

        elif isdict(email):  # emails labeled as 'home', 'work', etc
            for key, value in email.items():
                p.add_run(value)
                p.add_run('(' + key + ')')
                p.add_run(sep)

    path = os.path.join(resumecoder.path, path)
    data = yml_to_dict(path)

    name(data['name'], resumecoder.document)

    if 'email' in data:
        email(data['email'], resumecoder.document)


if __name__ == '__main__':
    rc = ResumeCoder('./data/Henry/')
    contact_info(rc, 'contact info.yml')
    rc.write()
