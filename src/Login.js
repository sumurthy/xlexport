var React = require('react');

////////////////////////////////////////////////////////////
class Login extends React.Component {
  state = {
    redirectToReferrer: false,
  }

  login = () => {
        console.log("Login pressed");
        this.setState({ redirectToReferrer: true })
        window.location.replace('https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=9b700ac9-4f07-4269-844f-afebf55c2dc2&response_type=token&redirect_uri=http://localhost:3000/auth/&scope=files.readwrite+offline_access&response_mode=fragment&state=12345&nonce=678910')
  }

  componentDidMount = () => {
      console.log("component did mount");
  }

  componentWillUnmount = () => {
      console.log("component will un-mount");
  }

  render() {
    console.log("render from login");
    console.log(this.context);
    return (
      <div>
          <p>
            You must log in to access Excel options
          </p>
        <button onClick={this.login}>Log in</button>
      </div>
    )
  }
}

module.exports = Login;
