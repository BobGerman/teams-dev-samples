import React from 'react';

/**
 * The web UI used when Teams pops out a browser window
 */
export default class WebPage extends React.Component {

  render() {
    return (
      <div>
        <h1>{process.env.REACT_APP_MANIFEST_NAME}</h1>
        <p>Your app is running in a stand-alone web page</p>
        <p>Your short message is not available</p>
      </div>
    );
  }

}
