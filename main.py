import word
import pathlib
import click
import preferences


@click.group()
@click.version_option(version="Release 1.0.2")
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
@click.argument("new_editing_time", type=click.IntRange(min=0, min_open=False, max=999999999, max_open=False))
def change_editing_time(file: pathlib.Path, new_editing_time: int):
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
    """Set all metadata fields to default values set in preferences.yaml"""
    if file.exists() is False:
        click.echo(click.style(f"File {file} was not found.", fg="red"))
        return
    if file.suffix == ".docx":
        word_file_metadata = word.Metadata(file)
        new_command_preferences = preferences.NewCommandPreferences()

        completed_with_errors = False

        editing_time = preferences.DEFAULT_WORD_EDITING_TIME
        try:
            editing_time = new_command_preferences.editing_time
        except AttributeError:
            click.secho('Preference for "editing time" not found. '
                        f'Default value "{preferences.DEFAULT_WORD_EDITING_TIME}" used.', fg="red")
            completed_with_errors = True
        except preferences.InvalidPreferencesStructureError:
            click.secho(f'Invalid structure of "{preferences.PREFERENCES_FILEPATH.name}" file.', fg="red")
            return
        except FileNotFoundError:
            click.secho(f'File "{preferences.PREFERENCES_FILEPATH.name}" not found.', fg="red")
            return
        except UnicodeDecodeError:
            click.secho(f'Encoding of "{preferences.PREFERENCES_FILEPATH.name}" must be UTF-8.', fg="red")
            return

        revision = preferences.DEFAULT_WORD_REVISION
        try:
            revision = new_command_preferences.revision
        except AttributeError:
            click.secho('Preference for "revision" not found. '
                        f'Default value "{preferences.DEFAULT_WORD_REVISION}" used.', fg="red")
            completed_with_errors = True
        except preferences.InvalidPreferencesStructureError:
            click.secho(f'Invalid structure of "{preferences.PREFERENCES_FILEPATH.name}" file.', fg="red")
            return
        except FileNotFoundError:
            click.secho(f'File "{preferences.PREFERENCES_FILEPATH.name}" not found.', fg="red")
            return
        except UnicodeDecodeError:
            click.secho(f'Encoding of "{preferences.PREFERENCES_FILEPATH.name}" must be UTF-8.', fg="red")
            return

        creator = preferences.DEFAULT_WORD_CREATOR
        try:
            creator = new_command_preferences.creator
        except AttributeError:
            click.secho('Preference for "creator" not found. '
                        f'Default value "{preferences.DEFAULT_WORD_CREATOR}" used.', fg="red")
            completed_with_errors = True
        except preferences.InvalidPreferencesStructureError:
            click.secho(f'Invalid structure of "{preferences.PREFERENCES_FILEPATH.name}" file.', fg="red")
            return
        except FileNotFoundError:
            click.secho(f'File "{preferences.PREFERENCES_FILEPATH.name}" not found.', fg="red")
            return
        except UnicodeDecodeError:
            click.secho(f'Encoding of "{preferences.PREFERENCES_FILEPATH.name}" must be UTF-8.', fg="red")
            return

        last_modified_by = preferences.DEFAULT_WORD_LAST_MODIFIED_BY
        try:
            last_modified_by = new_command_preferences.last_modified_by
        except AttributeError:
            click.secho('Preference for "last modified by" not found. '
                        f'Default value "{preferences.DEFAULT_WORD_LAST_MODIFIED_BY}" used.', fg="red")
            completed_with_errors = True
        except preferences.InvalidPreferencesStructureError:
            click.secho(f'Invalid structure of "{preferences.PREFERENCES_FILEPATH.name}" file.', fg="red")
            return
        except FileNotFoundError:
            click.secho(f'File "{preferences.PREFERENCES_FILEPATH.name}" not found.', fg="red")
            return
        except UnicodeDecodeError:
            click.secho(f'Encoding of "{preferences.PREFERENCES_FILEPATH.name}" must be UTF-8.', fg="red")
            return

        application_name = preferences.DEFAULT_WORD_APPLICATION_NAME
        try:
            application_name = new_command_preferences.application
        except AttributeError:
            click.secho('Preference for "application name" not found. '
                        f'Default value "{preferences.DEFAULT_WORD_APPLICATION_NAME}" used.', fg="red")
            completed_with_errors = True
        except preferences.InvalidPreferencesStructureError:
            click.secho(f'Invalid structure of "{preferences.PREFERENCES_FILEPATH.name}" file.', fg="red")
            return
        except FileNotFoundError:
            click.secho(f'File "{preferences.PREFERENCES_FILEPATH.name}" not found.', fg="red")
            return
        except UnicodeDecodeError:
            click.secho(f'Encoding of "{preferences.PREFERENCES_FILEPATH.name}" must be UTF-8.', fg="red")
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
    """Send hello to Smirnova according to preferences in preferences.yaml file"""
    if file.exists() is False:
        click.echo(click.style(f"File {file} was not found.", fg="red"))
        return
    if file.suffix == ".docx":
        word_file_metadata = word.Metadata(file)
        privet_smirnovoy_preferences = preferences.PrivetSmirnovoyPreference()

        completed_with_errors = False

        random_creators_string = None
        try:
            random_creators_string = privet_smirnovoy_preferences.random_creators_string
        except AttributeError:
            click.secho('Preference for "creator" not found.', fg="red")
            completed_with_errors = True
        except ValueError:
            click.secho('Preference "creators number" is invalid. Maybe too big?', fg="red")
            completed_with_errors = True
        except preferences.InvalidPreferencesStructureError:
            click.secho(f'Invalid structure of "{preferences.PREFERENCES_FILEPATH.name}" file.', fg="red")
            return
        except FileNotFoundError:
            click.secho(f'File "{preferences.PREFERENCES_FILEPATH.name}" not found.', fg="red")
            return
        except UnicodeDecodeError:
            click.secho(f'Encoding of "{preferences.PREFERENCES_FILEPATH.name}" must be UTF-8.', fg="red")
            return

        random_modifiers_string = None
        try:
            random_modifiers_string = privet_smirnovoy_preferences.random_modifiers_string
        except AttributeError:
            click.secho('Preference for "last modified by" not found.', fg="red")
            completed_with_errors = True
        except ValueError:
            click.secho('Preference "modifiers number" is invalid. Maybe too big?', fg="red")
            completed_with_errors = True
        except preferences.InvalidPreferencesStructureError:
            click.secho(f'Invalid structure of "{preferences.PREFERENCES_FILEPATH.name}" file.', fg="red")
            return
        except FileNotFoundError:
            click.secho(f'File "{preferences.PREFERENCES_FILEPATH.name}" not found.', fg="red")
            return
        except UnicodeDecodeError:
            click.secho(f'Encoding of "{preferences.PREFERENCES_FILEPATH.name}" must be UTF-8.', fg="red")
            return

        random_application = None
        try:
            random_application = privet_smirnovoy_preferences.random_application
        except AttributeError:
            click.secho('Preference for "application name" not found. ', fg="red")
            completed_with_errors = True
        except preferences.InvalidPreferencesStructureError:
            click.secho(f'Invalid structure of "{preferences.PREFERENCES_FILEPATH.name}" file.', fg="red")
            return
        except FileNotFoundError:
            click.secho(f'File "{preferences.PREFERENCES_FILEPATH.name}" not found.', fg="red")
            return
        except UnicodeDecodeError:
            click.secho(f'Encoding of "{preferences.PREFERENCES_FILEPATH.name}" must be UTF-8.', fg="red")
            return

        word_file_metadata.editing_time = preferences.PRIVET_SMIRNOVOY_EDITING_TIME
        word_file_metadata.revision = preferences.PRIVET_SMIRNOVOY_REVISION
        if random_creators_string is not None:
            word_file_metadata.creator = random_creators_string
        if random_modifiers_string is not None:
            word_file_metadata.last_modified_by = random_modifiers_string
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
