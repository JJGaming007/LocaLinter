// LocaLinterBridge.cs — drop this anywhere in your Unity project's Assets folder.
//
// It starts a tiny HTTP listener inside the running game that lets the LocaLinter
// agent read the exact strings the UI is displaying (with their on-screen rects,
// truncation state and overflow), take a screenshot, and click controls by id.
//
// Works identically in the Editor's Play Mode and in a development build on a
// device — on device the agent reaches it through `adb forward`.
//
// It compiles into the Editor and development builds only, so it can never ship
// in a release build. Nothing else in your project needs to change.
//
//   Port: 8791 by default. Override with the LOCALINTER_PORT environment
//   variable, or by editing DefaultPort below.

#if UNITY_EDITOR || DEVELOPMENT_BUILD

using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Text;
using System.Threading;
using UnityEngine;
using UnityEngine.EventSystems;
using UnityEngine.SceneManagement;
using UnityEngine.UI;

namespace LocaLinter
{
    public class LocaLinterBridge : MonoBehaviour
    {
        public const int DefaultPort = 8791;

        static LocaLinterBridge _instance;
        TcpListener _listener;
        Thread _thread;
        volatile bool _running;

        readonly object _queueLock = new object();
        readonly Queue<PendingRequest> _queue = new Queue<PendingRequest>();
        readonly Dictionary<int, GameObject> _elements = new Dictionary<int, GameObject>();

        [RuntimeInitializeOnLoadMethod(RuntimeInitializeLoadType.AfterSceneLoad)]
        static void Boot()
        {
            if (_instance != null) return;
            var go = new GameObject("~LocaLinterBridge");
            DontDestroyOnLoad(go);
            go.hideFlags = HideFlags.HideAndDontSave;
            _instance = go.AddComponent<LocaLinterBridge>();
        }

        void Start()
        {
            int port = DefaultPort;
            try
            {
                var env = Environment.GetEnvironmentVariable("LOCALINTER_PORT");
                if (!string.IsNullOrEmpty(env)) int.TryParse(env, out port);
            }
            catch { /* platform may deny env access */ }

            try
            {
                _listener = new TcpListener(IPAddress.Loopback, port);
                _listener.Start();
                _running = true;
                _thread = new Thread(AcceptLoop) { IsBackground = true, Name = "LocaLinterBridge" };
                _thread.Start();
                Debug.Log("[LocaLinter] Bridge listening on 127.0.0.1:" + port);
            }
            catch (Exception e)
            {
                Debug.LogWarning("[LocaLinter] Could not start the bridge on port " + port + ": " + e.Message);
            }
        }

        void OnDestroy() { Shutdown(); }
        void OnApplicationQuit() { Shutdown(); }

        void Shutdown()
        {
            _running = false;
            try { if (_listener != null) _listener.Stop(); } catch { }
            _listener = null;
        }

        // ── networking (background thread) ───────────────────────────────────

        void AcceptLoop()
        {
            while (_running)
            {
                TcpClient client = null;
                try
                {
                    client = _listener.AcceptTcpClient();
                }
                catch
                {
                    if (!_running) return;
                    continue;
                }
                var c = client;
                new Thread(() => Serve(c)) { IsBackground = true }.Start();
            }
        }

        void Serve(TcpClient client)
        {
            try
            {
                using (client)
                using (var stream = client.GetStream())
                {
                    client.ReceiveTimeout = 15000;
                    client.SendTimeout = 30000;

                    string method, path;
                    string body;
                    if (!ReadRequest(stream, out method, out path, out body)) return;

                    var pending = new PendingRequest
                    {
                        Method = method,
                        Path = path,
                        Body = body,
                        Done = new ManualResetEvent(false)
                    };
                    lock (_queueLock) _queue.Enqueue(pending);

                    // The main thread answers; 40s covers a slow end-of-frame capture.
                    if (!pending.Done.WaitOne(40000))
                    {
                        Write(stream, 504, "application/json", Encoding.UTF8.GetBytes("{\"error\":\"timed out on the main thread\"}"));
                        return;
                    }
                    Write(stream, pending.Status, pending.ContentType, pending.Payload);
                }
            }
            catch { /* a dropped client must never take the game down */ }
        }

        static bool ReadRequest(NetworkStream stream, out string method, out string path, out string body)
        {
            method = null; path = null; body = null;

            var head = new MemoryStream();
            var one = new byte[1];
            int matched = 0;
            while (matched < 4)
            {
                int n = stream.Read(one, 0, 1);
                if (n <= 0) return false;
                head.WriteByte(one[0]);
                if ((matched == 0 || matched == 2) && one[0] == (byte)'\r') matched++;
                else if ((matched == 1 || matched == 3) && one[0] == (byte)'\n') matched++;
                else matched = one[0] == (byte)'\r' ? 1 : 0;
                if (head.Length > 16384) return false;
            }

            var text = Encoding.UTF8.GetString(head.ToArray());
            var lines = text.Split(new[] { "\r\n" }, StringSplitOptions.None);
            var parts = lines[0].Split(' ');
            if (parts.Length < 2) return false;
            method = parts[0];
            path = parts[1];

            int length = 0;
            for (int i = 1; i < lines.Length; i++)
            {
                var l = lines[i];
                int colon = l.IndexOf(':');
                if (colon <= 0) continue;
                if (l.Substring(0, colon).Trim().ToLowerInvariant() == "content-length")
                    int.TryParse(l.Substring(colon + 1).Trim(), out length);
            }

            if (length > 0)
            {
                var buf = new byte[length];
                int read = 0;
                while (read < length)
                {
                    int n = stream.Read(buf, read, length - read);
                    if (n <= 0) break;
                    read += n;
                }
                body = Encoding.UTF8.GetString(buf, 0, read);
            }
            else body = "";

            return true;
        }

        static void Write(NetworkStream stream, int status, string contentType, byte[] payload)
        {
            var header = Encoding.UTF8.GetBytes(
                "HTTP/1.1 " + status + " " + (status == 200 ? "OK" : "Error") + "\r\n" +
                "Content-Type: " + contentType + "\r\n" +
                "Content-Length: " + payload.Length + "\r\n" +
                "Access-Control-Allow-Origin: *\r\n" +
                "Connection: close\r\n\r\n");
            stream.Write(header, 0, header.Length);
            stream.Write(payload, 0, payload.Length);
            stream.Flush();
        }

        // ── dispatch (main thread) ───────────────────────────────────────────

        void Update()
        {
            PendingRequest req = null;
            lock (_queueLock)
            {
                if (_queue.Count > 0) req = _queue.Dequeue();
            }
            if (req == null) return;

            try
            {
                Handle(req);
            }
            catch (Exception e)
            {
                req.Complete(500, "application/json", "{\"error\":" + Json.Str(e.Message) + "}");
            }
        }

        void Handle(PendingRequest req)
        {
            var path = req.Path;
            int q = path.IndexOf('?');
            if (q >= 0) path = path.Substring(0, q);

            switch (path)
            {
                case "/ping":
                    req.Complete(200, "application/json", Ping());
                    return;
                case "/state":
                    req.Complete(200, "application/json", State());
                    return;
                case "/screenshot":
                    StartCoroutine(Screenshot(req));
                    return;
                case "/tap":
                    req.Complete(200, "application/json", Tap(req.Body));
                    return;
                case "/click":
                    req.Complete(200, "application/json", Click(req.Body));
                    return;
                case "/longpress":
                    StartCoroutine(LongPress(req));
                    return;
                case "/back":
                    req.Complete(200, "application/json", Back());
                    return;
                case "/scroll":
                    req.Complete(200, "application/json", Scroll(req.Body));
                    return;
                case "/locale":
                    req.Complete(200, "application/json",
                        req.Method == "POST" ? SetLocale(req.Body) : GetLocale());
                    return;
                default:
                    req.Complete(404, "application/json", "{\"error\":\"unknown endpoint\"}");
                    return;
            }
        }

        // ── endpoints ────────────────────────────────────────────────────────

        string Ping()
        {
            var sb = new StringBuilder("{");
            Json.Bool(sb, "ok", true);
            Json.Comma(sb); Json.Str(sb, "mode", Application.isEditor ? "editor" : "player");
            Json.Comma(sb); Json.Str(sb, "product", Application.productName);
            Json.Comma(sb); Json.Str(sb, "unity", Application.unityVersion);
            Json.Comma(sb); Json.Str(sb, "platform", Application.platform.ToString());
            Json.Comma(sb); sb.Append("\"screen\":{\"width\":").Append(Screen.width)
                              .Append(",\"height\":").Append(Screen.height).Append("}");
            sb.Append("}");
            return sb.ToString();
        }

        string State()
        {
            _elements.Clear();

            var sb = new StringBuilder("{");
            Json.Str(sb, "scene", ActiveSceneNames());
            Json.Comma(sb); sb.Append("\"screen\":{\"width\":").Append(Screen.width)
                              .Append(",\"height\":").Append(Screen.height).Append("}");
            Json.Comma(sb); Json.Str(sb, "locale", CurrentLocaleCode());

            // ── texts ──
            Json.Comma(sb); sb.Append("\"texts\":[");
            bool first = true;
            foreach (var g in FindGraphics())
            {
                var info = TextInfo.From(g);
                if (info == null) continue;
                if (!first) sb.Append(',');
                first = false;
                info.Write(sb, this);
            }
            sb.Append(']');

            // ── interactables ──
            Json.Comma(sb); sb.Append("\"interactables\":[");
            first = true;
            foreach (var sel in FindActive<Selectable>())
            {
                if (!sel.interactable) continue;
                var go = sel.gameObject;
                _elements[go.GetInstanceID()] = go;
                if (!first) sb.Append(',');
                first = false;
                sb.Append('{');
                Json.Num(sb, "id", go.GetInstanceID());
                Json.Comma(sb); Json.Str(sb, "path", PathOf(go.transform));
                Json.Comma(sb); Json.Str(sb, "name", go.name);
                Json.Comma(sb); Json.Str(sb, "kind", sel.GetType().Name);
                Json.Comma(sb); Json.Str(sb, "label", LabelOf(go));
                Json.Comma(sb); WriteRect(sb, go.transform as RectTransform);
                sb.Append('}');
            }
            // Anything else that handles a click — custom cards, tiles, info badges.
            foreach (var h in FindActive<MonoBehaviour>())
            {
                if (!(h is IPointerClickHandler)) continue;
                var go = h.gameObject;
                if (_elements.ContainsKey(go.GetInstanceID())) continue;
                _elements[go.GetInstanceID()] = go;
                if (!first) sb.Append(',');
                first = false;
                sb.Append('{');
                Json.Num(sb, "id", go.GetInstanceID());
                Json.Comma(sb); Json.Str(sb, "path", PathOf(go.transform));
                Json.Comma(sb); Json.Str(sb, "name", go.name);
                Json.Comma(sb); Json.Str(sb, "kind", h.GetType().Name);
                Json.Comma(sb); Json.Str(sb, "label", LabelOf(go));
                Json.Comma(sb); WriteRect(sb, go.transform as RectTransform);
                sb.Append('}');
            }
            sb.Append(']');

            // ── scroll views ──
            Json.Comma(sb); sb.Append("\"scrolls\":[");
            first = true;
            foreach (var sr in FindActive<ScrollRect>())
            {
                var go = sr.gameObject;
                _elements[go.GetInstanceID()] = go;
                bool vertical = sr.vertical;
                float contentSize = 0f, viewSize = 0f;
                if (sr.content != null && sr.viewport != null)
                {
                    contentSize = vertical ? sr.content.rect.height : sr.content.rect.width;
                    viewSize = vertical ? sr.viewport.rect.height : sr.viewport.rect.width;
                }
                if (!first) sb.Append(',');
                first = false;
                sb.Append('{');
                Json.Num(sb, "id", go.GetInstanceID());
                Json.Comma(sb); Json.Str(sb, "path", PathOf(go.transform));
                Json.Comma(sb); Json.Bool(sb, "vertical", vertical);
                Json.Comma(sb); Json.Bool(sb, "canScroll", contentSize > viewSize + 4f);
                sb.Append('}');
            }
            sb.Append(']');

            sb.Append('}');
            return sb.ToString();
        }

        IEnumerator Screenshot(PendingRequest req)
        {
            yield return new WaitForEndOfFrame();
            Texture2D tex = null;
            try
            {
                tex = ScreenCapture.CaptureScreenshotAsTexture();
                var png = tex.EncodeToPNG();
                req.CompleteBinary(200, "image/png", png);
            }
            catch (Exception e)
            {
                req.Complete(500, "application/json", "{\"error\":" + Json.Str(e.Message) + "}");
            }
            finally
            {
                if (tex != null) Destroy(tex);
            }
        }

        string Tap(string body)
        {
            float x = Json.GetFloat(body, "x");
            float y = Json.GetFloat(body, "y");
            var go = RaycastAt(new Vector2(x, y));
            if (go == null) return "{\"ok\":false,\"reason\":\"nothing under that point\"}";
            ClickGameObject(go);
            return "{\"ok\":true,\"hit\":" + Json.Str(PathOf(go.transform)) + "}";
        }

        string Click(string body)
        {
            int id = (int)Json.GetFloat(body, "id");
            GameObject go;
            if (!_elements.TryGetValue(id, out go) || go == null)
                return "{\"ok\":false,\"reason\":\"element is gone — request /state again\"}";
            ClickGameObject(go);
            return "{\"ok\":true,\"hit\":" + Json.Str(PathOf(go.transform)) + "}";
        }

        IEnumerator LongPress(PendingRequest req)
        {
            float x = Json.GetFloat(req.Body, "x");
            float y = Json.GetFloat(req.Body, "y");
            float ms = Json.GetFloat(req.Body, "ms");
            if (ms <= 0) ms = 800;

            var go = RaycastAt(new Vector2(x, y));
            if (go == null)
            {
                req.Complete(200, "application/json", "{\"ok\":false,\"reason\":\"nothing under that point\"}");
                yield break;
            }
            var data = PointerData(new Vector2(x, y), go);
            ExecuteEvents.Execute(go, data, ExecuteEvents.pointerEnterHandler);
            ExecuteEvents.Execute(go, data, ExecuteEvents.pointerDownHandler);
            yield return new WaitForSecondsRealtime(ms / 1000f);
            ExecuteEvents.Execute(go, data, ExecuteEvents.pointerUpHandler);
            req.Complete(200, "application/json", "{\"ok\":true,\"hit\":" + Json.Str(PathOf(go.transform)) + "}");
        }

        // Unity has no OS back button, so this clicks whatever looks like one.
        // On Android the agent falls back to the hardware key, which Unity
        // delivers as Escape, so this only matters in the Editor.
        string Back()
        {
            string[] words = { "back", "close", "cancel", "return", "×", "x", "✕" };
            foreach (var sel in FindActive<Selectable>())
            {
                if (!sel.interactable) continue;
                var label = (LabelOf(sel.gameObject) + " " + sel.gameObject.name).ToLowerInvariant();
                foreach (var w in words)
                {
                    if (label.Contains(w))
                    {
                        ClickGameObject(sel.gameObject);
                        return "{\"ok\":true,\"via\":" + Json.Str(PathOf(sel.transform)) + "}";
                    }
                }
            }
            return "{\"ok\":false,\"reason\":\"no back or close control on this screen\"}";
        }

        string Scroll(string body)
        {
            int id = (int)Json.GetFloat(body, "id");
            float pos = Json.GetFloat(body, "position");
            GameObject go;
            if (!_elements.TryGetValue(id, out go) || go == null)
                return "{\"ok\":false,\"reason\":\"scroll view is gone\"}";
            var sr = go.GetComponent<ScrollRect>();
            if (sr == null) return "{\"ok\":false,\"reason\":\"not a scroll view\"}";
            pos = Mathf.Clamp01(pos);
            if (sr.vertical) sr.verticalNormalizedPosition = pos;
            else sr.horizontalNormalizedPosition = pos;
            Canvas.ForceUpdateCanvases();
            return "{\"ok\":true}";
        }

        string GetLocale()
        {
            var code = CurrentLocaleCode();
            return "{\"code\":" + Json.Str(code) + ",\"name\":" + Json.Str(code) + "}";
        }

        string SetLocale(string body)
        {
            var code = Json.GetString(body, "code");
            var settings = Type.GetType("UnityEngine.Localization.Settings.LocalizationSettings, Unity.Localization");
            if (settings == null)
                return "{\"ok\":false,\"reason\":\"Unity Localization package is not installed — switch the language in-game\"}";
            try
            {
                var availableProp = settings.GetProperty("AvailableLocales", BindingFlags.Public | BindingFlags.Static);
                var available = availableProp.GetValue(null, null);
                var locales = available.GetType().GetProperty("Locales").GetValue(available, null) as IEnumerable;
                foreach (var loc in locales)
                {
                    var id = loc.GetType().GetProperty("Identifier").GetValue(loc, null);
                    var codeStr = id.GetType().GetProperty("Code").GetValue(id, null) as string;
                    if (string.Equals(codeStr, code, StringComparison.OrdinalIgnoreCase))
                    {
                        settings.GetProperty("SelectedLocale", BindingFlags.Public | BindingFlags.Static)
                                .SetValue(null, loc, null);
                        return "{\"ok\":true,\"code\":" + Json.Str(codeStr) + "}";
                    }
                }
                return "{\"ok\":false,\"reason\":\"locale not available in this build\"}";
            }
            catch (Exception e)
            {
                return "{\"ok\":false,\"reason\":" + Json.Str(e.Message) + "}";
            }
        }

        // ── Unity helpers ────────────────────────────────────────────────────

        static string ActiveSceneNames()
        {
            var sb = new StringBuilder();
            for (int i = 0; i < SceneManager.sceneCount; i++)
            {
                var s = SceneManager.GetSceneAt(i);
                if (!s.isLoaded) continue;
                if (sb.Length > 0) sb.Append('+');
                sb.Append(s.name);
            }
            return sb.ToString();
        }

        static string CurrentLocaleCode()
        {
            var settings = Type.GetType("UnityEngine.Localization.Settings.LocalizationSettings, Unity.Localization");
            if (settings != null)
            {
                try
                {
                    var sel = settings.GetProperty("SelectedLocale", BindingFlags.Public | BindingFlags.Static)
                                      .GetValue(null, null);
                    if (sel != null)
                    {
                        var id = sel.GetType().GetProperty("Identifier").GetValue(sel, null);
                        return id.GetType().GetProperty("Code").GetValue(id, null) as string;
                    }
                }
                catch { }
            }
            return Application.systemLanguage.ToString();
        }

        static IEnumerable<T> FindActive<T>() where T : Component
        {
            var all = Resources.FindObjectsOfTypeAll<T>();
            foreach (var c in all)
            {
                if (c == null) continue;
                var go = c.gameObject;
                if (!go.activeInHierarchy) continue;
                if (!go.scene.IsValid()) continue;              // prefab / asset, not on screen
                if ((go.hideFlags & HideFlags.HideInHierarchy) != 0) continue;
                var b = c as Behaviour;
                if (b != null && !b.enabled) continue;
                yield return c;
            }
        }

        static IEnumerable<Graphic> FindGraphics()
        {
            foreach (var g in FindActive<Graphic>())
            {
                if (g.canvas == null) continue;
                if (!g.canvas.enabled) continue;
                if (EffectiveAlpha(g) < 0.05f) continue;
                yield return g;
            }
        }

        static float EffectiveAlpha(Graphic g)
        {
            float a = g.color.a;
            var t = g.transform;
            while (t != null)
            {
                var cg = t.GetComponent<CanvasGroup>();
                if (cg != null)
                {
                    a *= cg.alpha;
                    if (cg.ignoreParentGroups) break;
                }
                t = t.parent;
            }
            return a;
        }

        public static string PathOf(Transform t)
        {
            var sb = new StringBuilder(t.name);
            var p = t.parent;
            int guard = 0;
            while (p != null && guard++ < 24)
            {
                sb.Insert(0, p.name + "/");
                p = p.parent;
            }
            return sb.ToString();
        }

        static string LabelOf(GameObject go)
        {
            foreach (var g in go.GetComponentsInChildren<Graphic>(false))
            {
                var s = TextInfo.ReadText(g);
                if (!string.IsNullOrEmpty(s) && s.Trim().Length > 0) return s.Trim();
            }
            return go.name;
        }

        void WriteRect(StringBuilder sb, RectTransform rt)
        {
            sb.Append("\"rect\":");
            if (rt == null) { sb.Append("null"); return; }
            var r = ScreenRect(rt);
            sb.Append("{\"x\":").Append(F(r.x))
              .Append(",\"y\":").Append(F(r.y))
              .Append(",\"w\":").Append(F(r.width))
              .Append(",\"h\":").Append(F(r.height)).Append('}');
        }

        /// Screen-space rect with the origin at the TOP-LEFT, matching how the
        /// screenshot and every reader of this data thinks about coordinates.
        public static Rect ScreenRect(RectTransform rt)
        {
            var corners = new Vector3[4];
            rt.GetWorldCorners(corners);
            var canvas = rt.GetComponentInParent<Canvas>();
            Camera cam = null;
            if (canvas != null && canvas.renderMode != RenderMode.ScreenSpaceOverlay) cam = canvas.worldCamera;

            float minX = float.MaxValue, minY = float.MaxValue, maxX = float.MinValue, maxY = float.MinValue;
            for (int i = 0; i < 4; i++)
            {
                var p = RectTransformUtility.WorldToScreenPoint(cam, corners[i]);
                minX = Mathf.Min(minX, p.x); maxX = Mathf.Max(maxX, p.x);
                minY = Mathf.Min(minY, p.y); maxY = Mathf.Max(maxY, p.y);
            }
            return new Rect(minX, Screen.height - maxY, maxX - minX, maxY - minY);
        }

        static string F(float v)
        {
            return v.ToString("0.##", CultureInfo.InvariantCulture);
        }

        GameObject RaycastAt(Vector2 screenPos)
        {
            if (EventSystem.current == null) return null;
            var data = new PointerEventData(EventSystem.current) { position = screenPos };
            var hits = new List<RaycastResult>();
            EventSystem.current.RaycastAll(data, hits);
            foreach (var h in hits)
            {
                if (h.gameObject == null) continue;
                var handler = ExecuteEvents.GetEventHandler<IPointerClickHandler>(h.gameObject);
                if (handler != null) return handler;
            }
            return hits.Count > 0 ? hits[0].gameObject : null;
        }

        PointerEventData PointerData(Vector2 pos, GameObject go)
        {
            return new PointerEventData(EventSystem.current)
            {
                position = pos,
                button = PointerEventData.InputButton.Left,
                clickCount = 1,
                pointerPress = go,
                pointerCurrentRaycast = new RaycastResult { gameObject = go, screenPosition = pos }
            };
        }

        void ClickGameObject(GameObject go)
        {
            var rt = go.transform as RectTransform;
            var centre = rt != null ? ScreenRect(rt).center : new Vector2(Screen.width / 2f, Screen.height / 2f);
            var pos = new Vector2(centre.x, Screen.height - centre.y);
            var data = PointerData(pos, go);

            var target = ExecuteEvents.GetEventHandler<IPointerClickHandler>(go) ?? go;
            ExecuteEvents.Execute(target, data, ExecuteEvents.pointerEnterHandler);
            ExecuteEvents.Execute(target, data, ExecuteEvents.pointerDownHandler);
            ExecuteEvents.Execute(target, data, ExecuteEvents.pointerUpHandler);
            ExecuteEvents.Execute(target, data, ExecuteEvents.pointerClickHandler);
            ExecuteEvents.Execute(target, data, ExecuteEvents.submitHandler);
        }

        // ── text extraction (uGUI Text and TextMeshPro, via reflection) ──────

        class TextInfo
        {
            static readonly Dictionary<Type, Members> Cache = new Dictionary<Type, Members>();

            class Members
            {
                public PropertyInfo Text, PreferredWidth, PreferredHeight, Truncated, FontSize, Font, HasOverflow;
                public bool IsTextComponent;
            }

            static Members For(Type t)
            {
                Members m;
                if (Cache.TryGetValue(t, out m)) return m;
                m = new Members();
                m.Text = t.GetProperty("text", typeof(string));
                m.PreferredWidth = t.GetProperty("preferredWidth", typeof(float));
                m.PreferredHeight = t.GetProperty("preferredHeight", typeof(float));
                m.Truncated = t.GetProperty("isTextTruncated", typeof(bool));
                m.HasOverflow = t.GetProperty("isTextOverflowing", typeof(bool));
                m.FontSize = t.GetProperty("fontSize");
                m.Font = t.GetProperty("font") ?? t.GetProperty("fontSharedMaterial");
                m.IsTextComponent = m.Text != null;
                Cache[t] = m;
                return m;
            }

            public static string ReadText(Graphic g)
            {
                if (g == null) return null;
                var m = For(g.GetType());
                if (!m.IsTextComponent) return null;
                try { return m.Text.GetValue(g, null) as string; } catch { return null; }
            }

            public Graphic Graphic;
            public string Text;
            public float PreferredWidth = -1, PreferredHeight = -1;
            public bool Truncated;
            public string FontName = "";
            public float FontSize;

            public static TextInfo From(Graphic g)
            {
                var m = For(g.GetType());
                if (!m.IsTextComponent) return null;
                var info = new TextInfo { Graphic = g };
                try { info.Text = m.Text.GetValue(g, null) as string; } catch { return null; }
                if (info.Text == null) info.Text = "";

                try { if (m.PreferredWidth != null) info.PreferredWidth = (float)m.PreferredWidth.GetValue(g, null); } catch { }
                try { if (m.PreferredHeight != null) info.PreferredHeight = (float)m.PreferredHeight.GetValue(g, null); } catch { }
                try { if (m.Truncated != null) info.Truncated = (bool)m.Truncated.GetValue(g, null); } catch { }
                if (!info.Truncated)
                {
                    try { if (m.HasOverflow != null) info.Truncated = (bool)m.HasOverflow.GetValue(g, null); } catch { }
                }
                try
                {
                    if (m.FontSize != null)
                    {
                        var v = m.FontSize.GetValue(g, null);
                        info.FontSize = v is float ? (float)v : Convert.ToSingle(v);
                    }
                }
                catch { }
                try
                {
                    if (m.Font != null)
                    {
                        var f = m.Font.GetValue(g, null) as UnityEngine.Object;
                        if (f != null) info.FontName = f.name;
                    }
                }
                catch { }
                return info;
            }

            public void Write(StringBuilder sb, LocaLinterBridge bridge)
            {
                var go = Graphic.gameObject;
                sb.Append('{');
                Json.Num(sb, "id", go.GetInstanceID());
                Json.Comma(sb); Json.Str(sb, "path", PathOf(go.transform));
                Json.Comma(sb); Json.Str(sb, "text", Text);
                Json.Comma(sb); Json.Str(sb, "component", Graphic.GetType().Name);
                Json.Comma(sb); Json.Bool(sb, "active", true);
                Json.Comma(sb); Json.Bool(sb, "isTruncated", Truncated);
                if (PreferredWidth >= 0) { Json.Comma(sb); Json.Num(sb, "preferredWidth", PreferredWidth); }
                if (PreferredHeight >= 0) { Json.Comma(sb); Json.Num(sb, "preferredHeight", PreferredHeight); }
                Json.Comma(sb); Json.Num(sb, "fontSize", FontSize);
                Json.Comma(sb); Json.Str(sb, "font", FontName);
                Json.Comma(sb); bridge.WriteRect(sb, Graphic.rectTransform);
                sb.Append('}');
            }
        }

        // ── request plumbing ─────────────────────────────────────────────────

        class PendingRequest
        {
            public string Method, Path, Body;
            public ManualResetEvent Done;
            public int Status = 200;
            public string ContentType = "application/json";
            public byte[] Payload = new byte[0];

            public void Complete(int status, string contentType, string body)
            {
                Status = status;
                ContentType = contentType + "; charset=utf-8";
                Payload = Encoding.UTF8.GetBytes(body ?? "");
                Done.Set();
            }

            public void CompleteBinary(int status, string contentType, byte[] body)
            {
                Status = status;
                ContentType = contentType;
                Payload = body ?? new byte[0];
                Done.Set();
            }
        }

        // ── minimal JSON writing / reading ───────────────────────────────────

        static class Json
        {
            public static void Comma(StringBuilder sb) { sb.Append(','); }

            public static void Str(StringBuilder sb, string key, string value)
            {
                sb.Append('"').Append(key).Append("\":").Append(Str(value));
            }

            public static void Num(StringBuilder sb, string key, float value)
            {
                sb.Append('"').Append(key).Append("\":").Append(value.ToString("0.####", CultureInfo.InvariantCulture));
            }

            public static void Bool(StringBuilder sb, string key, bool value)
            {
                sb.Append('"').Append(key).Append("\":").Append(value ? "true" : "false");
            }

            public static string Str(string s)
            {
                if (s == null) return "null";
                var sb = new StringBuilder(s.Length + 8);
                sb.Append('"');
                foreach (var c in s)
                {
                    switch (c)
                    {
                        case '"': sb.Append("\\\""); break;
                        case '\\': sb.Append("\\\\"); break;
                        case '\n': sb.Append("\\n"); break;
                        case '\r': sb.Append("\\r"); break;
                        case '\t': sb.Append("\\t"); break;
                        default:
                            if (c < 0x20 || c == 0x7f) sb.Append("\\u").Append(((int)c).ToString("x4"));
                            else sb.Append(c);
                            break;
                    }
                }
                sb.Append('"');
                return sb.ToString();
            }

            /// Good enough for the flat {"x":1,"y":2} bodies this bridge receives.
            public static float GetFloat(string body, string key)
            {
                var s = Raw(body, key);
                float v;
                return float.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out v) ? v : 0f;
            }

            public static string GetString(string body, string key)
            {
                var s = Raw(body, key);
                if (s == null) return "";
                s = s.Trim();
                if (s.Length >= 2 && s[0] == '"') s = s.Substring(1, s.Length - 2);
                return s.Replace("\\\"", "\"").Replace("\\\\", "\\");
            }

            static string Raw(string body, string key)
            {
                if (string.IsNullOrEmpty(body)) return null;
                var needle = "\"" + key + "\"";
                int i = body.IndexOf(needle, StringComparison.Ordinal);
                if (i < 0) return null;
                i = body.IndexOf(':', i + needle.Length);
                if (i < 0) return null;
                i++;
                int start = i;
                bool inString = false;
                while (i < body.Length)
                {
                    char c = body[i];
                    if (c == '"' && (i == 0 || body[i - 1] != '\\')) inString = !inString;
                    else if (!inString && (c == ',' || c == '}')) break;
                    i++;
                }
                return body.Substring(start, i - start).Trim();
            }
        }
    }
}

#endif
