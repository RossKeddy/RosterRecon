import pytest
import rosterrecon
from openpyxl import load_workbook
from rosterrecon import NameMutator

# Test name mutations

TEST_NAMES = {
    1: "John Smith",
    2: "John Davidson-Smith",
    3: "John-Paul Smith-Robinson",
    4: "José Gonzáles",
    5: "🙂 Emoji Folks 🙂",
    6: "Jean-Charles Martin",
    7: "Madonna Wayne Gacey",   # 3-word name: exercises the 'second' field in mutations
}


def test_rate_limit_is_in_milliseconds(monkeypatch):
    monkeypatch.setattr(
        'sys.argv',
        ['rosterrecon', '--company', 'example', '--rate-limit', '750'],
    )

    args = rosterrecon.parse_arguments()

    assert args.rate_limit_ms == 750


def test_rate_limit_rejects_negative_values(monkeypatch):
    monkeypatch.setattr(
        'sys.argv',
        ['rosterrecon', '--company', 'example', '--rate-limit', '-1'],
    )

    with pytest.raises(SystemExit):
        rosterrecon.parse_arguments()


def test_rate_limit_converts_milliseconds_to_seconds(monkeypatch):
    sleep_calls = []
    monkeypatch.setattr(rosterrecon.time, 'sleep', sleep_calls.append)

    rosterrecon.apply_rate_limit(750)
    rosterrecon.apply_rate_limit(0)

    assert sleep_calls == [0.75]


def test_f_last():
    name = TEST_NAMES[1]
    mutator = NameMutator(name)
    assert mutator.f_last() == set(["jsmith"])

    # Hyphenated last: compound form + each part
    name = TEST_NAMES[2]
    mutator = NameMutator(name)
    assert mutator.f_last() == set(["jdavidson-smith", "jdavidson", "jsmith"])

    # Hyphenated first and last: compound last + each last part; first stays compound
    name = TEST_NAMES[3]
    mutator = NameMutator(name)
    assert mutator.f_last() == set(["jsmith-robinson", "jsmith", "jrobinson"])

    name = TEST_NAMES[4]
    mutator = NameMutator(name)
    assert mutator.f_last() == set(["jgonzales"])

    name = TEST_NAMES[5]
    mutator = NameMutator(name)
    assert mutator.f_last() == set(["efolks"])

    # Compound hyphenated first name: jean-charles.martin must be generated (issue #82)
    name = TEST_NAMES[6]
    mutator = NameMutator(name)
    assert mutator.f_last() == set(["jmartin"])

    # 3-word name: second field produces an extra variant using the middle name
    name = TEST_NAMES[7]
    mutator = NameMutator(name)
    assert mutator.f_last() == set(["mgacey", "mwayne"])


def test_f_dot_last():
    name = TEST_NAMES[1]
    mutator = NameMutator(name)
    assert mutator.f_dot_last() == set(["j.smith"])

    name = TEST_NAMES[2]
    mutator = NameMutator(name)
    assert mutator.f_dot_last() == set(["j.davidson-smith", "j.davidson", "j.smith"])

    name = TEST_NAMES[3]
    mutator = NameMutator(name)
    assert mutator.f_dot_last() == set(["j.smith-robinson", "j.smith", "j.robinson"])

    name = TEST_NAMES[4]
    mutator = NameMutator(name)
    assert mutator.f_dot_last() == set(["j.gonzales"])

    name = TEST_NAMES[5]
    mutator = NameMutator(name)
    assert mutator.f_dot_last() == set(["e.folks"])

    name = TEST_NAMES[6]
    mutator = NameMutator(name)
    assert mutator.f_dot_last() == set(["j.martin"])

    name = TEST_NAMES[7]
    mutator = NameMutator(name)
    assert mutator.f_dot_last() == set(["m.gacey", "m.wayne"])


def test_last_f():
    name = TEST_NAMES[1]
    mutator = NameMutator(name)
    assert mutator.last_f() == set(["smithj"])

    name = TEST_NAMES[2]
    mutator = NameMutator(name)
    assert mutator.last_f() == set(["davidson-smithj", "davidsonj", "smithj"])

    name = TEST_NAMES[3]
    mutator = NameMutator(name)
    assert mutator.last_f() == set(["smith-robinsonj", "smithj", "robinsonj"])

    name = TEST_NAMES[4]
    mutator = NameMutator(name)
    assert mutator.last_f() == set(["gonzalesj"])

    name = TEST_NAMES[5]
    mutator = NameMutator(name)
    assert mutator.last_f() == set(["folkse"])

    name = TEST_NAMES[6]
    mutator = NameMutator(name)
    assert mutator.last_f() == set(["martinj"])

    name = TEST_NAMES[7]
    mutator = NameMutator(name)
    assert mutator.last_f() == set(["gaceym", "waynem"])


def test_first_dot_last():
    name = TEST_NAMES[1]
    mutator = NameMutator(name)
    assert mutator.first_dot_last() == set(["john.smith"])

    name = TEST_NAMES[2]
    mutator = NameMutator(name)
    assert mutator.first_dot_last() == set(["john.davidson-smith", "john.davidson", "john.smith"])

    # Compound first name is preserved intact; last name variants are expanded
    name = TEST_NAMES[3]
    mutator = NameMutator(name)
    assert mutator.first_dot_last() == set(["john-paul.smith-robinson", "john-paul.smith", "john-paul.robinson"])

    name = TEST_NAMES[4]
    mutator = NameMutator(name)
    assert mutator.first_dot_last() == set(["jose.gonzales"])

    name = TEST_NAMES[5]
    mutator = NameMutator(name)
    assert mutator.first_dot_last() == set(["emoji.folks"])

    # The core fix for issue #82: compound first name generates the correct username
    name = TEST_NAMES[6]
    mutator = NameMutator(name)
    assert mutator.first_dot_last() == set(["jean-charles.martin"])

    name = TEST_NAMES[7]
    mutator = NameMutator(name)
    assert mutator.first_dot_last() == set(["madonna.gacey", "madonna.wayne"])


def test_first_l():
    name = TEST_NAMES[1]
    mutator = NameMutator(name)
    assert mutator.first_l() == set(["johns"])

    # davidson-smith[0]='d', davidson[0]='d' (dup), smith[0]='s'
    name = TEST_NAMES[2]
    mutator = NameMutator(name)
    assert mutator.first_l() == set(["johnd", "johns"])

    # smith-robinson[0]='s', smith[0]='s' (dup), robinson[0]='r'
    name = TEST_NAMES[3]
    mutator = NameMutator(name)
    assert mutator.first_l() == set(["john-pauls", "john-paulr"])

    name = TEST_NAMES[4]
    mutator = NameMutator(name)
    assert mutator.first_l() == set(["joseg"])

    name = TEST_NAMES[5]
    mutator = NameMutator(name)
    assert mutator.first_l() == set(["emojif"])

    name = TEST_NAMES[6]
    mutator = NameMutator(name)
    assert mutator.first_l() == set(["jean-charlesm"])

    name = TEST_NAMES[7]
    mutator = NameMutator(name)
    assert mutator.first_l() == set(["madonnag", "madonnaw"])


def test_first():
    name = TEST_NAMES[1]
    mutator = NameMutator(name)
    assert mutator.first() == set(["john"])

    name = TEST_NAMES[2]
    mutator = NameMutator(name)
    assert mutator.first() == set(["john"])

    # Compound first name is preserved intact
    name = TEST_NAMES[3]
    mutator = NameMutator(name)
    assert mutator.first() == set(["john-paul"])

    name = TEST_NAMES[4]
    mutator = NameMutator(name)
    assert mutator.first() == set(["jose"])

    name = TEST_NAMES[5]
    mutator = NameMutator(name)
    assert mutator.first() == set(["emoji"])

    name = TEST_NAMES[6]
    mutator = NameMutator(name)
    assert mutator.first() == set(["jean-charles"])

    # 3-word name: first() always returns only the first token regardless of second/last
    name = TEST_NAMES[7]
    mutator = NameMutator(name)
    assert mutator.first() == set(["madonna"])


def test_hyphen_variants():
    assert NameMutator._hyphen_variants("smith") == ["smith"]
    assert NameMutator._hyphen_variants("davidson-smith") == ["davidson-smith", "davidson", "smith"]
    assert NameMutator._hyphen_variants("a-b-c") == ["a-b-c", "a", "b", "c"]


def test_clean_name():
    mutator = NameMutator("xxx")
    assert mutator.clean_name("  🙂Ànèôõö    ßï🙂  ") == "aneooo ssi"

    name = "Dr. Hannibal Lecter, PhD."
    assert mutator.clean_name(name) == "hannibal lecter"

    name = "Mr. Fancy Pants MD, PhD, MBA"
    assert mutator.clean_name(name) == "fancy pants"

    name = "Mr. Cert Dude (OSCP, OSCE)"
    assert mutator.clean_name(name) == "cert dude"


def test_split_name():
    mutator = NameMutator("xxx")

    name = "madonna wayne gacey"
    assert mutator.split_name(name) == {"first": "madonna", "second": "wayne", "last": "gacey"}

    name = "twiggy ramirez"
    assert mutator.split_name(name) == {"first": "twiggy", "second": "", "last": "ramirez"}

    name = "brian warner is marilyn manson"
    assert mutator.split_name(name) == {"first": "brian", "second": "marilyn", "last": "manson"}

    # Hyphens within a name segment are preserved (not treated as word separators)
    name = "jean-charles martin"
    assert mutator.split_name(name) == {"first": "jean-charles", "second": "", "last": "martin"}

    name = "john davidson-smith"
    assert mutator.split_name(name) == {"first": "john", "second": "", "last": "davidson-smith"}

    name = "john-paul smith-robinson"
    assert mutator.split_name(name) == {"first": "john-paul", "second": "", "last": "smith-robinson"}


def test_find_employees():
    with open("tests/mock-employee-response", "r") as infile:
        result = infile.read()
    employees = rosterrecon.find_employees(result)

    assert len(employees) == 2
    assert employees[0] == {
        'full_name': 'Michael Myers',
        'job_title': 'Camp Counsellor',
        'location': 'xxxxx',
    }
    assert employees[1] == {
        'full_name': 'Freddy Krueger',
        'job_title': 'Babysitter',
        'location': 'xxxxx',
    }

    with open("tests/mock-employee-response-last-page", "r") as infile:
        result = infile.read()
    assert not rosterrecon.find_employees(result)


def test_deduplicate_employees():
    employees = [
        {'full_name': 'Jane Doe', 'job_title': 'Engineer', 'location': 'Toronto'},
        {'full_name': 'jane doe', 'job_title': 'engineer', 'location': 'toronto'},
        {'full_name': 'Jane Doe', 'job_title': 'Manager', 'location': 'Toronto'},
    ]

    result = rosterrecon.deduplicate_employees(employees)

    assert len(result) == 2
    assert result[0]['job_title'] == 'Engineer'
    assert result[1]['job_title'] == 'Manager'


def test_write_files_creates_csv_and_xlsx(tmp_path):
    employees = [
        {
            'full_name': 'Jane Doe',
            'job_title': 'Director, Security',
            'location': 'Toronto, Canada',
        },
        {
            'full_name': 'John Smith',
            'job_title': '=HYPERLINK("https://example.test")',
            'location': 'London',
        },
    ]

    rosterrecon.write_files('example', '@example.com', employees, tmp_path)

    csv_path = tmp_path / 'example-employees.csv'
    assert csv_path.exists()
    with csv_path.open(newline='', encoding='utf-8') as infile:
        rows = list(__import__('csv').DictReader(infile))
    assert rows[0] == {
        'full_name': 'Jane Doe',
        'job_title': 'Director, Security',
        'location': 'Toronto, Canada',
    }
    assert rows[1]['job_title'].startswith("'=")

    workbook = load_workbook(tmp_path / 'example-employees.xlsx', data_only=False)
    assert workbook.sheetnames == ['Employees', 'Summary']
    worksheet = workbook['Employees']
    assert [cell.value for cell in worksheet[1]][:3] == [
        'Full Name', 'Job Title', 'Location'
    ]
    assert worksheet['B2'].value == 'Director, Security'
    assert worksheet['D2'].value == 'j.doe@example.com'
    assert worksheet['B3'].data_type != 'f'
    assert worksheet.auto_filter.ref == 'A1:I3'
    assert worksheet.freeze_panes == 'A2'

    summary = workbook['Summary']
    assert summary['B3'].value == 2
    assert summary['B5'].value == 2
    assert summary['D7'].value == 'Top Job Titles'
    assert summary['D8'].value == 'Director, Security'


def test_output_stem_prevents_path_traversal(tmp_path):
    employee = {
        'full_name': 'Jane Doe',
        'job_title': 'Engineer',
        'location': 'Toronto',
    }

    rosterrecon.write_files('../../target co', '', [employee], tmp_path)

    assert (tmp_path / 'target_co-employees.xlsx').exists()
    assert not (tmp_path.parent / 'target co-employees.xlsx').exists()

