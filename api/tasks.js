export default async function handler(req, res) {
  const owner = process.env.GITHUB_OWNER;
  const repo = process.env.GITHUB_REPO;
  const token = process.env.GITHUB_TOKEN;
  const path = "tasks.json";

  const url = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;

  async function getFile() {
    const r = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/vnd.github+json"
      }
    });

    if (!r.ok) {
      throw new Error(`GitHub read failed: ${r.status}`);
    }

    const data = await r.json();
    const content = JSON.parse(
      Buffer.from(data.content, "base64").toString("utf8")
    );

    return { content, sha: data.sha };
  }

  if (req.method === "GET") {
    try {
      const { content } = await getFile();
      return res.status(200).json(content);
    } catch (e) {
      return res.status(500).json({ error: e.message });
    }
  }

  if (req.method === "PUT") {
    try {
      const { sha } = await getFile();

      const newContent = {
        tasks: req.body.tasks || [],
        meetings: req.body.meetings || [],
        reminders: req.body.reminders || [],
        aiHistory: req.body.aiHistory || []
      };

      const update = await fetch(url, {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/vnd.github+json",
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          message: "Update tasks.json from MyTaskManager",
          content: Buffer.from(
            JSON.stringify(newContent, null, 2)
          ).toString("base64"),
          sha
        })
      });

      if (!update.ok) {
        const err = await update.json();
        throw new Error(err.message || `GitHub update failed: ${update.status}`);
      }

      return res.status(200).json(newContent);
    } catch (e) {
      return res.status(500).json({ error: e.message });
    }
  }

  return res.status(405).json({ error: "Method not allowed" });
}
