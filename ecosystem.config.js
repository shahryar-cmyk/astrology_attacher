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
      repo: "git@github.com:shahryar-cmyk/astrology_attacher.git",
      path: "/home/ubuntu/astrology_attacher",
      "pre-deploy-local": "",
      "post-deploy":
        "pip install -r requirements.txt && pm2 reload ecosystem.config.js --env production",
      "pre-setup": "",
      ssh_options: "ForwardAgent=yes",
    },
  },
};
