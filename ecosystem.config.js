module.exports = {
  apps: [
    {
      name: "my-python-app", // Name of the Python app
      script: "venv/bin/python", // Path to the Python interpreter
      args: "app.py", // The main script of your Python application
      watch: true, // Enable watching
      watch: ["."], // Watch the current directory
    },
    {
      name: "service-worker", // Name of the service worker
      script: "./service-worker/index.js", // Main script of the service worker
      watch: ["./service-worker"], // Watch the service worker directory
    },
  ],

  deploy: {
    production: {
      key: "Afaq_New_Server_Key.pem",
      user: "ubuntu",
      host: "3.107.68.178",
      ref: "origin/main",
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
