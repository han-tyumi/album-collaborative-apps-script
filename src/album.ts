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
   * @param spotifyUri The URI for the album on Spotify.
   * @param spotifyUrl The URL to the album on Spotify.
   */
  constructor(
    public title: string = '',
    public artist: string = '',
    public submitter: string = '',
    public spotifyUri: string = '',
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
        'artist' | 'title' | 'submitter' | 'spotifyUri' | 'spotifyUrl'
      >,
    ): boolean => {
      const response = ui.prompt('New Album', prompt, ui.ButtonSet.OK_CANCEL);
      if (response.getSelectedButton() === ui.Button.OK) {
        this[property] = response.getResponseText();
        return true;
      }
      return false;
    };

    if (submitter) {
      this.submitter = submitter;
    } else if (
      !prompt("Enter the name of the album's submitter.", 'submitter')
    ) {
      return false;
    }

    if (prompt("Enter the album's Spotify URI.", 'spotifyUri')) {
      // Get album ID from Spotify URI.
      const albumId = this.spotifyUri.split(':').pop();
      const clientSecret = PropertiesService.getScriptProperties().getProperty(
        'SpotifyClientSecret',
      );

      if (albumId && clientSecret) {
        // Fetch our Spotify access token.
        const accessToken = JSON.parse(
          UrlFetchApp.fetch('https://accounts.spotify.com/api/token', {
            method: 'post',
            payload: {
              grant_type: 'client_credentials',
            },
            headers: {
              Authorization: `Basic ${Utilities.base64EncodeWebSafe(
                clientSecret,
              )}`,
            },
          }).getContentText(),
        ).access_token;

        // Fetch the Spotify album object.
        const albumObject = JSON.parse(
          UrlFetchApp.fetch(`https://api.spotify.com/v1/albums/${albumId}`, {
            headers: {
              Authorization: `Bearer ${accessToken}`,
            },
          }).getContentText(),
        );

        // Set the title, artist, and URL using the album object.
        this.title = albumObject.name;
        this.artist = albumObject.artists[0].name;
        this.spotifyUrl = albumObject.external_urls.spotify;

        return true;
      }
      ui.alert('Unable to fetch album data.');
    }

    if (!prompt("Enter the album's title.", 'title')) {
      return false;
    }

    if (!prompt("Enter the album's artist.", 'artist')) {
      return false;
    }

    if (!prompt("Enter the album's Spotify URL.", 'spotifyUrl')) {
      return false;
    }

    return true;
  }
}
