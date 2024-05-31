module.exports = {
  apps: [
    {
      name: "astrology_attacher", // Name of the Python app
      script: "venv/bin/python", // Path to the Python interpreter
      args: "app.py", // The main script of your Python application
      watch: true, // Enable watching
      watch: ["."], // Watch the current directory
    },
  ],

  deploy: {
    production: {
      key: "elcaminoquecreas.pem",
      user: "ubuntu",
      host: "18.226.181.19",
      ref: "origin/main",
      repo: "git@github.com:shahryar-cmyk/astrology_attacher.git",
      path: "/home/ubuntu/astrology_attacher", // Correct path where the project will be deployed
      "pre-deploy-local": "",
      "post-deploy": "cd /home/ubuntu/astrology_attacher && ./deploy.sh",
      "pre-setup": "",
      ssh_options: "ForwardAgent=yes",
    },
  },
};
