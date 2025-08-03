Sure! Here is the complete **Installation Guide for running n8n without Docker Compose**, using only `docker run`, specifically for **Windows Server 2019** with **Docker CE (Moby)** installed using the official Microsoft script.

---

## 🛠️ **Installation Guide: Running n8n Locally Using Docker CE (No Compose)**

**Windows Server 2019 – Modern Method**

---

### 1. Install Docker CE (Moby) Using Microsoft’s Official Script

1. Open **PowerShell as Administrator**
2. Download the Docker installation script:

```powershell
Invoke-WebRequest -UseBasicParsing `
  "https://raw.githubusercontent.com/microsoft/Windows-Containers/Main/helpful_tools/Install-DockerCE/install-docker-ce.ps1" `
  -OutFile install-docker-ce.ps1

```

1. Run the script:

```powershell
.\install-docker-ce.ps1

```

> The script will:
> 
> - Enable the Containers feature
> - Install Docker CE (Moby)
> - Start the Docker service
1. Restart the server manually or with:

```powershell
Restart-Computer -Force

```

1. After reboot, confirm Docker is installed:

```powershell
docker version

```

---

### 2. Run n8n Using `docker run` (No Compose)

Once Docker is ready, run n8n directly with:

```powershell
docker run -it --name n8n-test `
  -p 5678:5678 `
  -e N8N_BASIC_AUTH_ACTIVE=true `
  -e N8N_BASIC_AUTH_USER=admin `
  -e N8N_BASIC_AUTH_PASSWORD=123 `
  n8nio/n8n

```

> This starts n8n in interactive mode and binds it to http://localhost:5678
> 

---

### 3. Access n8n Interface

Open your browser and go to:

```
http://localhost:5678

```

- **Username:** `admin`
- **Password:** `123`

---

### 4. Optional: Run n8n in Background (Detached Mode)

To keep n8n running in the background:

```powershell
docker run -d --name n8n `
  -p 5678:5678 `
  -e N8N_BASIC_AUTH_ACTIVE=true `
  -e N8N_BASIC_AUTH_USER=admin `
  -e N8N_BASIC_AUTH_PASSWORD=123 `
  n8nio/n8n

```

---

### 5. Manage the Container

- **Stop the container:**

```powershell
docker stop n8n

```

- **Start it again:**

```powershell
docker start n8n

```

- **View logs:**

```powershell
docker logs n8n

```

- **Remove it:**

```powershell
docker rm -f n8n

```

---

### 🔍 Testing Outcome

If your system does **not** support Linux containers, running the above will result in:

```
exec format error

```

or

```
image operating system "linux" cannot be used on this platform

```

This confirms that **Windows Server 2019 only supports Windows containers**, and not Linux containers like the `n8nio/n8n` image.

---

Let me know if you want the same steps with environment variables pulled from a `.env` file or for use in a scheduled task or service.
