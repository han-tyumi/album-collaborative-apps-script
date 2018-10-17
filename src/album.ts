/**
 * Represents an album to be reviewed.
 * @todo Add a property to track the form linked to the album.
 */
export class Album {
  /**
   * Initializes a new album.
   * @param title The title of the album.
   * @param artist The artist of the album.
   * @param submitter The name of the submitter of the album.
   * @param spotifyUrl The URL to the album on Spotify.
   */
  constructor(
    public title: string = '',
    public artist: string = '',
    public submitter: string = '',
    public spotifyUrl: string = '',
  ) {}

  /**
   * A formatted string representing the album.
   */
  public get formattedName(): string {
    return `${this.title} — ${this.artist}`;
  }

  /**
   * Prompts the user for the album's information.
   * @param submitter The name of the submitter of the album.
   * @returns Whether or not all data was retrieved.
   */
  public prompt(submitter?: string): boolean {
    const ui = SpreadsheetApp.getUi();

    const prompt = (
      prompt: string,
      property: keyof Pick<
        Album,
        'artist' | 'title' | 'submitter' | 'spotifyUrl'
      >,
    ): boolean => {
      const response = ui.prompt('New Album', prompt, ui.ButtonSet.OK_CANCEL);
      if (response.getSelectedButton() === ui.Button.OK) {
        this[property] = response.getResponseText();
        return true;
      }
      return false;
    };

    if (!prompt("Enter the album's title.", 'title')) {
      return false;
    }

    if (!prompt("Enter the album's artist.", 'artist')) {
      return false;
    }

    if (submitter) {
      this.submitter = submitter;
    } else if (
      !prompt("Enter the name of the album's submitter.", 'submitter')
    ) {
      return false;
    }

    if (!prompt("Enter the album's Spotify URL.", 'spotifyUrl')) {
      return false;
    }

    return true;
  }
}
