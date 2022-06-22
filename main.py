import word
import pathlib
import click
import preferences


# Constants
PREFERENCES_FILEPATH = pathlib.Path("preferences.yaml")
# Command "new"
DEFAULT_WORD_EDITING_TIME = 0
DEFAULT_WORD_REVISION = 1
DEFAULT_WORD_CREATOR = "admin"
DEFAULT_WORD_LAST_MODIFIED_BY = "admin"
DEFAULT_WORD_APPLICATION_NAME = "Microsoft Office Word"
# Command "privet_smirnovoy"
PRIVET_SMIRNOVOY_EDITING_TIME = 599940
PRIVET_SMIRNOVOY_REVISION = 9999999


@click.group()
@click.version_option(version="Beta 1.2")
def main():
    pass


@click.command()
@click.argument("file", type=pathlib.Path)  # File or path of the next document types: .docx
def get_metadata(file: pathlib.Path):
    """Get file known metadata"""
    if file.exists() is False:
        click.echo(click.style(f"File {file} was not found.", fg="red"))
        return

    if file.suffix == ".docx":
        word_file_metadata = word.Metadata(file)
        click.echo(click.style("Creator: ", fg="yellow") +
                   click.style(word_file_metadata.creator, fg="white"))
        click.echo(click.style("Last modified by: ", fg="yellow") +
                   click.style(word_file_metadata.last_modified_by, fg="white"))
        click.echo(click.style("Revision: ", fg="yellow") +
                   click.style(word_file_metadata.revision, fg="white"))
        click.echo(click.style("Application: ", fg="yellow") +
                   click.style(word_file_metadata.application_name, fg="white"))
        click.echo(click.style("Editing time: ", fg="yellow") +
                   click.style(word_file_metadata.editing_time, fg="white"))
    else:
        click.echo(click.style(f"File type {file.suffix} is not yet available.", fg="red"))


@click.command()
@click.argument("file", type=pathlib.Path)
@click.argument("new_creator", type=str)
def change_creator(file: pathlib.Path, new_creator: str):
    """Change file creator"""
    if file.exists() is False:
        click.echo(click.style(f"File {file} was not found.", fg="red"))
        return

    if file.suffix == ".docx":
        word_file_metadata = word.Metadata(file)
        word_file_metadata.creator = new_creator
        click.secho("Success.", fg="green")
    else:
        click.echo(click.style(f"File type {file.suffix} is not yet available.", fg="red"))


@click.command()
@click.argument("file", type=pathlib.Path)
@click.argument("new_modifier", type=str)
def change_modifier(file: pathlib.Path, new_modifier: str):
    """Change file modifier"""
    if file.exists() is False:
        click.echo(click.style(f"File {file} was not found.", fg="red"))
        return

    if file.suffix == ".docx":
        word_file_metadata = word.Metadata(file)
        word_file_metadata.last_modified_by = new_modifier
        click.secho("Success.", fg="green")
    else:
        click.echo(click.style(f"File type {file.suffix} is not yet available.", fg="red"))


@click.command()
@click.argument("file", type=pathlib.Path)
@click.argument("new_revision", type=int)
def change_revision(file: pathlib.Path, new_revision: int):
    """Change file revision"""
    if file.exists() is False:
        click.echo(click.style(f"File {file} was not found.", fg="red"))
        return

    if file.suffix == ".docx":
        word_file_metadata = word.Metadata(file)
        word_file_metadata.revision = new_revision
        click.secho("Success.", fg="green")
    else:
        click.echo(click.style(f"File type {file.suffix} is not yet available.", fg="red"))


@click.command()
@click.argument("file", type=pathlib.Path)
@click.argument("new_application", type=str)
def change_application(file: pathlib.Path, new_application: str):
    """Change file application"""
    if file.exists() is False:
        click.echo(click.style(f"File {file} was not found.", fg="red"))
        return

    if file.suffix == ".docx":
        word_file_metadata = word.Metadata(file)
        word_file_metadata.application_name = new_application
        click.secho("Success.", fg="green")
    else:
        click.echo(click.style(f"File type {file.suffix} is not yet available.", fg="red"))


@click.command()
@click.argument("file", type=pathlib.Path)
@click.argument("new_editing_time", type=int)
def change_editing_time(file: pathlib.Path, new_editing_time: str):
    """Change file editing time"""
    if file.exists() is False:
        click.echo(click.style(f"File {file} was not found.", fg="red"))
        return

    if file.suffix == ".docx":
        word_file_metadata = word.Metadata(file)
        word_file_metadata.editing_time = new_editing_time
        click.secho("Success.", fg="green")
    else:
        click.echo(click.style(f"File type {file.suffix} is not yet available.", fg="red"))


@click.command()
@click.argument("file", type=pathlib.Path)
def new(file: pathlib.Path):
    if file.exists() is False:
        click.echo(click.style(f"File {file} was not found.", fg="red"))
        return
    if file.suffix == ".docx":
        word_file_metadata = word.Metadata(file)
        new_command_preferences = preferences.NewCommandPreferences(PREFERENCES_FILEPATH)

        completed_with_errors = False

        editing_time = DEFAULT_WORD_EDITING_TIME
        try:
            editing_time = new_command_preferences.editing_time
        except preferences.PreferenceNotFoundError:
            click.secho('Preference for "editing time" not found. '
                        f'Default value "{DEFAULT_WORD_EDITING_TIME}" used.', fg="red")
            completed_with_errors = True
        except FileNotFoundError:
            click.secho(f'File "{PREFERENCES_FILEPATH.name}" not found.', fg="red")
            return

        revision = DEFAULT_WORD_REVISION
        try:
            revision = new_command_preferences.revision
        except preferences.PreferenceNotFoundError:
            click.secho('Preference for "revision" not found. '
                        f'Default value "{DEFAULT_WORD_REVISION}" used.', fg="red")
            completed_with_errors = True
        except FileNotFoundError:
            click.secho(f'File "{PREFERENCES_FILEPATH.name}" not found.', fg="red")
            return

        creator = DEFAULT_WORD_CREATOR
        try:
            creator = new_command_preferences.creator
        except preferences.PreferenceNotFoundError:
            click.secho('Preference for "creator" not found. '
                        f'Default value "{DEFAULT_WORD_CREATOR}" used.', fg="red")
            completed_with_errors = True
        except FileNotFoundError:
            click.secho(f'File "{PREFERENCES_FILEPATH.name}" not found.', fg="red")
            return

        last_modified_by = DEFAULT_WORD_LAST_MODIFIED_BY
        try:
            last_modified_by = new_command_preferences.last_modified_by
        except preferences.PreferenceNotFoundError:
            click.secho('Preference for "last modified by" not found. '
                        f'Default value "{DEFAULT_WORD_LAST_MODIFIED_BY}" used.', fg="red")
            completed_with_errors = True
        except FileNotFoundError:
            click.secho(f'File "{PREFERENCES_FILEPATH.name}" not found.', fg="red")
            return

        application_name = DEFAULT_WORD_APPLICATION_NAME
        try:
            application_name = new_command_preferences.application
        except preferences.PreferenceNotFoundError:
            click.secho('Preference for "application name" not found. '
                        f'Default value "{DEFAULT_WORD_APPLICATION_NAME}" used.', fg="red")
            completed_with_errors = True
        except FileNotFoundError:
            click.secho(f'File "{PREFERENCES_FILEPATH.name}" not found.', fg="red")
            return

        word_file_metadata.editing_time = editing_time
        word_file_metadata.revision = revision
        word_file_metadata.creator = creator
        word_file_metadata.last_modified_by = last_modified_by
        word_file_metadata.application_name = application_name

        if completed_with_errors:
            click.secho("Completed with errors.", fg="yellow")
            click.secho("Please, check preferences.yaml", fg="yellow")
        else:
            click.secho("Success.", fg="green")
    else:
        click.echo(click.style(f"File type {file.suffix} is not yet available.", fg="red"))


@click.command()
@click.argument("file", type=pathlib.Path)
def privet_smirnovoy(file: pathlib.Path):
    if file.exists() is False:
        click.echo(click.style(f"File {file} was not found.", fg="red"))
        return
    if file.suffix == ".docx":
        word_file_metadata = word.Metadata(file)
        privet_smirnovoy_preferences = preferences.PrivetSmirnovoyPreference(PREFERENCES_FILEPATH)

        completed_with_errors = False

        random_creators_string = None
        try:
            random_creators_string = privet_smirnovoy_preferences.random_creators_string
        except preferences.PreferenceNotFoundError:
            click.secho('Preference for "creator" not found.', fg="red")
            completed_with_errors = True
        except preferences.InvalidPreferenceValueError:
            click.secho('Preference "creators number" is invalid. Maybe too big?', fg="red")
            completed_with_errors = True
        except FileNotFoundError:
            click.secho(f'File "{PREFERENCES_FILEPATH.name}" not found.', fg="red")
            return

        random_modifiers_string = None
        try:
            random_modifiers_string = privet_smirnovoy_preferences.random_modifiers_string
        except preferences.PreferenceNotFoundError:
            click.secho('Preference for "last modified by" not found.', fg="red")
            completed_with_errors = True
        except preferences.InvalidPreferenceValueError:
            click.secho('Preference "modifiers number" is invalid. Maybe too big?', fg="red")
            completed_with_errors = True
        except FileNotFoundError:
            click.secho(f'File "{PREFERENCES_FILEPATH.name}" not found.', fg="red")
            return

        random_application = None
        try:
            random_application = privet_smirnovoy_preferences.random_application
        except preferences.PreferenceNotFoundError:
            click.secho('Preference for "application name" not found. ', fg="red")
            completed_with_errors = True
        except FileNotFoundError:
            click.secho(f'File "{PREFERENCES_FILEPATH.name}" not found.', fg="red")
            return

        word_file_metadata.editing_time = PRIVET_SMIRNOVOY_EDITING_TIME
        word_file_metadata.revision = PRIVET_SMIRNOVOY_REVISION
        if random_creators_string is not None:
            word_file_metadata.creator = random_creators_string
        if random_modifiers_string is not None:
            word_file_metadata.last_modified_by = privet_smirnovoy_preferences.random_modifiers_string
        if random_application is not None:
            word_file_metadata.application_name = random_application

        if completed_with_errors:
            click.secho("Completed with errors.", fg="yellow")
            click.secho("Please, check preferences.yaml", fg="yellow")
        else:
            click.secho("Success.", fg="green")
    else:
        click.echo(click.style(f"File type {file.suffix} is not yet available.", fg="red"))


main.add_command(get_metadata)
main.add_command(change_creator)
main.add_command(change_modifier)
main.add_command(change_revision)
main.add_command(change_application)
main.add_command(change_editing_time)
main.add_command(new)
main.add_command(privet_smirnovoy)


if __name__ == "__main__":
    main()
