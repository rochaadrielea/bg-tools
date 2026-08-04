# Beyond Gravity tools - server deployment

Two tools, one command:
- **adab**   - ADAB Compare + Find Batch web app  -> http://chbs4212:8000
- **tracker** - As-Built tracker (static page)      -> http://chbs4212:8080

Everything runs in Docker so it is the same on every machine, and Git keeps
the history of changes.

---

## One-time setup on the server (chbs4212)

### 0. Make a workspace you own
Your login has no home folder yet, so create one:

    sudo mkdir -p "$HOME"
    sudo chown -R "$(id -u):$(id -g)" "$HOME"
    cd "$HOME" && pwd

### 1. Install Docker (you have sudo)

    sudo apt update
    sudo apt install -y docker.io docker-compose-v2 git unzip
    sudo usermod -aG docker "$(whoami)"      # run docker without sudo

Log out and back in once (or run `newgrp docker`) so the group takes effect.

### 2. Tell Docker about the corporate proxy (so it can pull images)

    sudo mkdir -p /etc/systemd/system/docker.service.d
    sudo tee /etc/systemd/system/docker.service.d/http-proxy.conf >/dev/null <<'EOF'
    [Service]
    Environment="HTTP_PROXY=http://chbs8055.ruaggroup.com:8080"
    Environment="HTTPS_PROXY=http://chbs8055.ruaggroup.com:8080"
    Environment="NO_PROXY=localhost,127.0.0.1"
    EOF
    sudo systemctl daemon-reload
    sudo systemctl restart docker

Test that Docker can reach the internet through the proxy:

    docker run --rm hello-world

If you see "Hello from Docker!" the proxy works. If it hangs, IT has not
opened Docker Hub through the proxy - tell me and we switch to the no-Docker path.

---

## Put the code on the server

Drag `bgtools.zip` into your MobaXterm file panel (left side), then:

    cd "$HOME"
    unzip bgtools.zip        # one command - not manual
    cd bgtools

## Start version control (Git)

    git init
    git add .
    git commit -m "First version of my tools"

(Later: create a PRIVATE repo on GitHub, then
`git remote add origin <url>` and `git push`.)

## Run everything

    docker compose up -d --build

Open in a browser on the network:
- ADAB:    http://chbs4212:8000
- Tracker: http://chbs4212:8080

## Update after you change a file

    git add . && git commit -m "what I changed"
    docker compose up -d --build

## Handy commands

    docker compose ps           # what is running
    docker compose logs -f adab # watch the ADAB log
    docker compose down         # stop everything
