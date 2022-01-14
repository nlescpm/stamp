#!/usr/bin/env python
import argparse
import curses
import json
import pathlib
from Levenshtein import distance
import pandas as pd
import npyscreen
from mailmerge import MailMerge

parser = argparse.ArgumentParser(description='docx/xlsx Merge tool.')
parser.add_argument('templates', type=pathlib.Path, nargs='+',
                    help='DOCX templates to merge')

group = parser.add_mutually_exclusive_group(required=True)
group.add_argument('-x', '--xlsx', type=pathlib.Path, help="XLSX file")
group.add_argument('-c', '--csv',  type=pathlib.Path, help="CSV file")

ENTRY_LABEL_OFFSET = 30

template_by = None
key_mapping = {}
template_mapping = {}


class MyTestApp(npyscreen.NPSAppManaged):
    """Application wrapper"""

    def onStart(self):
        self.registerForm("MAIN", MainForm())
        self.registerForm("TEMPLATES", TemplatesForm())
        self.registerForm("OUTPUT", OutputForm())


class MainForm(npyscreen.Form):
    """This form that will be presented to the user."""

    def create(self):
        """Create the form."""
        self.name = "Map template on the data"
        self.add(
                npyscreen.TitleText,
                begin_entry_at=ENTRY_LABEL_OFFSET,
                editable=False,
                name="Data file",
                value=source_name
                )
        self.add(
                npyscreen.TitleCombo,
                begin_entry_at=ENTRY_LABEL_OFFSET,
                maxlen=1,
                name="Select template by",
                values=list(source_df.keys()),
                scroll_exit=False,
                )
        self.nextrely += 1

        for field in fields:
            # find the edit distance to each column
            distances = [distance(field, column) for column in source_df.columns]

            self.add(
                npyscreen.TitleCombo,
                begin_entry_at=ENTRY_LABEL_OFFSET,
                maxlen=1,
                name=field,
                value=distances.index(min(distances)),
                values=list(source_df.columns),
                scroll_exit=False,
            )
        self.add_handlers({curses.KEY_DC: self.wipe_value})

    def wipe_value(self, _):
        wipe_widget = self.get_widget(self.editw)
        wipe_widget.set_value(None)

    def afterEditing(self):
        """When 'OK' is pressed."""
        template_by_widget = self.get_widget(1)
        template_by_idx = template_by_widget.get_value()
        if template_by_idx is not None:
            global template_by
            template_by = template_by_widget.get_values()[template_by_idx]

            # move to the next form
            self.parentApp.setNextForm("TEMPLATES")
        else:
            self.parentApp.setNextForm(None)


class TemplatesForm(npyscreen.Form):
    """This form that will be presented to the user."""

    def create(self):
        """Create the form."""
        self.name = "Template selection"
        self.add_handlers({curses.KEY_DC: self.wipe_value})

    def beforeEditing(self):
        global template_by
        # find the possible values the template_by key takes
        values_to_map = list(set(source_df[template_by]))
        values_to_map.sort(key=lambda x: str(x))

        self.add(
                npyscreen.TitleText,
                begin_entry_at=ENTRY_LABEL_OFFSET,
                editable=False,
                name=template_by,
                value="Template file"
                )

        for name in values_to_map:
            # find the edit distance to each template
            distances = [distance(name, t.name) for t in args.templates]

            self.add(
                npyscreen.TitleCombo,
                maxlen=1,
                begin_entry_at=ENTRY_LABEL_OFFSET,
                name=str(name),
                value=distances.index(min(distances)),
                values=[t.name for t in args.templates],
                scroll_exit=False,
            )

    def afterEditing(self):
        """When 'OK' is pressed."""
        for widget in self._widgets__[1:]:
            idx = widget.get_value()
            if idx is not None:
                value = widget.get_values()[idx]
                template_mapping[widget.name] = value
            else:
                template_mapping[widget.name] = None

        self.parentApp.setNextForm("OUTPUT")

    def wipe_value(self, _):
        widget = self.get_widget(self.editw)
        widget.set_value(None)


class OutputForm(npyscreen.Form):
    """This form that will be presented to the user."""

    def create(self):
        """Create the form."""
        self.name = "Output"
        self.add_handlers({curses.KEY_DC: self.wipe_value})

    def wipe_value(self, _):
        widget = self.get_widget(self.editw)
        widget.set_value(None)

    def afterEditing(self):
        """When 'OK' is pressed."""
        self.parentApp.setNextForm(None)


if __name__ == "__main__":
    args = parser.parse_args()

    if args.xlsx:
        source_df = pd.read_excel(args.xlsx)
        source_name = str(args.xlsx)
    elif args.csv:
        source_df = pd.read_csv(args.csv)
        source_name = str(args.csv)

    # column 'BCO ' => 'ACCEPT', 'REJECT'
    sfields = set()

    for template in args.templates:
        mm = MailMerge(template)
        mf = mm.get_merge_fields()
        for f in mf:
            sfields.add(f)

    fields = list(sfields)
    fields.sort()

    TA = MyTestApp()
    TA.run()

    for idx, field in enumerate(fields):
        widget = TA.getForm("MAIN").get_widget(idx + 2)

        # make sure we have the right widget
        assert field == widget.name

        v = widget.get_value()
        if v is not None:
            key_mapping[field] = widget.get_values()[v]
        else:
            key_mapping[field] = None

    # TODO: print JSON format of mapping
    print(json.dumps(key_mapping))
    print(json.dumps(template_mapping))

    # doc=  MailMerge('./CIT_2021_Letter-Final-Accept.docx'),

    # # mapping from merge_fields to excel column:
    # mapping = {
    #     'Dossiernummer':     'Dossiernummer',
    #     'Kenmerk_NLeSC':     'Kenmerk NLeSC',
    #     'Title':             'Title',
    #     'Main_applicant':    'Main applicant',
    #     'Affiliation_lvl_1': 'Affiliation (lvl 1)',
    #     'Affiliation_lvl_2': 'Affiliation (lvl 2)',
    #     'Affiliation_lvl_3': 'Affiliation (lvl 3)',
    #     'Address':           'Address',
    #     'Zip_CodeCity':      'Zip Code/City',
    #     'Motivation1':       'Motivation1',
    #     'Motivation2':       'Motivation2',
    #     'Motivation3':       'Motivation3',
    #     'TotaalFTE':         'TotaalFTE',
    #     'WaardeFTE':         'WaardeFTE'
    # }


    # for row in range(rows):
    #     try:
    #         document = templates[df['BCO '][row]]
    #         fields = document.get_merge_fields()
    #         args = {}
    #         for field in fields:
    #             args[field] = df[mapping[field]][row]
    #         document.merge(**args)
    #         document.write(df['Kenmerk NLeSC'][row] + '.docx')
    #         print('Kenmerk:', df['Kenmerk NLeSC'][row], df['BCO '][row])
    #     except:
    #         print('no template for:', df['BCO '][row])
