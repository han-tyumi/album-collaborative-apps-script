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
   * @param spotifyUri The URL to the album on Spotify.
   */
  constructor(
    public title: string = '',
    public artist: string = '',
    public submitter: string = '',
    public spotifyUri: string = '',
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
        'artist' | 'title' | 'submitter' | 'spotifyUri'
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

    if (!prompt("Enter the album's Spotify URI.", 'spotifyUri')) {
      return false;
    }

    // Fetch album details.
    const albumId = this.spotifyUri.split(':').pop();

    if (albumId) {
      const accessToken = JSON.parse(
        UrlFetchApp.fetch('https://accounts.spotify.com/api/token', {
          method: 'post',
          payload: {
            grant_type: 'client_credentials',
          },
          headers: {
            Authorization: `Basic ${Utilities.base64EncodeWebSafe(
              'e90cdca00ec946508bb13ea20d1afb16:265dd27d98be4a8db484cd33ff7f7783',
            )}`,
          },
        }).getContentText(),
      ).access_token;

      const response = UrlFetchApp.fetch(
        `https://api.spotify.com/v1/albums/${albumId}`,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        },
      );

      ui.alert(response.getContentText());
    }

    return true;
  }
}
