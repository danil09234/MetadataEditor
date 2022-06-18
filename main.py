import word
import pathlib
import click


@click.group()
@click.version_option(version="Beta 1.0")
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


main.add_command(get_metadata)
main.add_command(change_creator)
main.add_command(change_modifier)
main.add_command(change_revision)


if __name__ == "__main__":
    main()
