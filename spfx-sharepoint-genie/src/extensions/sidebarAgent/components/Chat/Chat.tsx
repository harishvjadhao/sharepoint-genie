import * as React from "react";
import { useState, useRef, useEffect, useCallback } from "react";
// import { AzureOpenAI } from "openai";
// @ts-ignore - types may not be present in SPFx environment, suppress for build
import ReactMarkdown from "react-markdown";
// @ts-ignore
import remarkGfm from "remark-gfm";
import styles from "./Chat.module.scss";
import { initSession } from "../../services/initSession";
import { chat } from "../../services/chat";
import { IUser } from "../../models/ISidebarAgentProperties";

export interface IChatProps {
  showTyping?: boolean;
  currentUserLogin: IUser;
  baseUrl: string;
}

interface Message {
  role: "user" | "assistant";
  content: string;
}

interface AttachmentMeta {
  id: string;
  name: string;
  size: number;
  type: string;
  text?: string;
}

const Chat: React.FC<IChatProps> = ({
  showTyping = true,
  currentUserLogin,
  baseUrl,
}) => {
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string>();
  const [attachments, setAttachments] = useState<AttachmentMeta[]>([]);
  const [sessionId, setSessionId] = useState<string>();
  const [sessionInitializing, setSessionInitializing] = useState(false);

  const listRef = useRef<HTMLDivElement | null>(null);
  const abortRef = useRef<AbortController | null>(null);
  const textAreaRef = useRef<HTMLTextAreaElement | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const scrollToBottom = useCallback(() => {
    if (listRef.current) {
      listRef.current.scrollTop = listRef.current.scrollHeight;
    }
  }, []);

  useEffect(() => {
    scrollToBottom();
  }, [messages, scrollToBottom, loading]);

  useEffect(() => {
    let isMounted = true;
    setSessionInitializing(true);

    const init = async () => {
      try {
        const session = await initSession(baseUrl, currentUserLogin);
        if (!session && isMounted) {
          setError("Session not initialized.");
          setSessionInitializing(false);
          return;
        }
        if (isMounted) setSessionId(session.sessionId);
      } catch (err) {
        if (isMounted) setError("Session init failed.");
      } finally {
        if (isMounted) setSessionInitializing(false);
      }
    };

    init();

    return () => {
      isMounted = false;
    };
  }, []);

  const handleSend = async () => {
    if (!input.trim() || loading) return;

    if (!sessionId) {
      setError("Session not initialized.");
      return;
    }

    setError(undefined);
    // const attachmentText = attachments
    //   .filter((a) => a.text)
    //   .map((a) => `\n\n[Attachment: ${a.name}]\n${a.text}`)
    //   .join("");
    const userMessage: Message = {
      role: "user",
      // content: input.trim() + attachmentText,
      content: input.trim(),
    };
    setMessages((prev) => [...prev, userMessage]);
    setInput("");
    setAttachments([]);
    // Return focus immediately so user can start typing the next prompt.
    requestAnimationFrame(() => textAreaRef.current?.focus());
    setLoading(true);

    abortRef.current?.abort();
    abortRef.current = new AbortController();

    try {
      const response = await chat(
        sessionId,
        input.trim(),
        currentUserLogin.loginName
      );
      console.log("Response from custom chat service:", response);

      const content = response.reply || "No response from service.";
      const botMessage: Message = { role: "assistant", content };
      setMessages((prev) => [...prev, botMessage]);
    } catch (e: any) {
      if (e?.name === "AbortError") return; // silently ignore
      setError("Error processing your request.");
    } finally {
      setLoading(false);
      // Ensure focus restored after response (covers screen reader / other focus shifts)
      requestAnimationFrame(() => textAreaRef.current?.focus());
    }
  };

  const onSelectFiles = async (fileList: FileList | null) => {
    if (!fileList) return;
    const maxFiles = 5;
    const maxSize = 200 * 1024; // 200KB
    const supported = ["text/plain", "application/json", "text/markdown"];
    const collected: AttachmentMeta[] = [];
    for (let i = 0; i < fileList.length && collected.length < maxFiles; i++) {
      const f = fileList[i];
      const meta: AttachmentMeta = {
        id: `${Date.now()}-${i}-${f.name}`,
        name: f.name,
        size: f.size,
        type: f.type,
      };
      if (
        f.size <= maxSize &&
        (supported.indexOf(f.type) !== -1 || /\.(txt|md|json)$/i.test(f.name))
      ) {
        try {
          meta.text = (await f.text()).slice(0, 8000);
        } catch {
          /* ignore */
        }
      }
      collected.push(meta);
    }
    setAttachments((prev) => [...prev, ...collected]);
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  const removeAttachment = (id: string) => {
    setAttachments((prev) => prev.filter((a) => a.id !== id));
  };

  const handleKeyDown: React.KeyboardEventHandler<HTMLTextAreaElement> = (
    e
  ) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  const handleCopy = (text: string) => {
    void navigator.clipboard.writeText(text);
  };

  const renderMessage = (msg: Message, idx: number) => {
    const isUser = msg.role === "user";
    return (
      <div
        key={idx}
        className={`${styles.messageRow} ${isUser ? styles.userRow : ""}`}
      >
        <div
          className={`${styles.avatar} ${
            !isUser ? styles.assistantAvatar : ""
          }`}
          aria-hidden="true"
        >
          {isUser
            ? currentUserLogin.loginName?.[0]?.toUpperCase() || "U"
            : "Ge"}
        </div>
        <div className={`${styles.bubble} ${isUser ? styles.userBubble : ""}`}>
          {isUser ? (
            msg.content
          ) : (
            <ReactMarkdown
              // @ts-ignore
              remarkPlugins={[remarkGfm]}
              linkTarget="_blank"
              components={{
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                a: (props: any) => <a {...props} rel="noreferrer" />,
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                code: ({ inline, className, children, ...rest }: any) => {
                  const txt = String(children);
                  if (inline) {
                    return <code {...rest}>{txt}</code>;
                  }
                  return (
                    <pre className={styles.codeBlock}>
                      <code {...rest}>{txt}</code>
                    </pre>
                  );
                },
              }}
            >
              {msg.content}
            </ReactMarkdown>
          )}
          {!isUser && (
            <div className={styles.bubbleActions}>
              <button
                type="button"
                className={styles.iconButton}
                aria-label="Copy message"
                onClick={() => handleCopy(msg.content)}
              >
                Copy
              </button>
            </div>
          )}
        </div>
      </div>
    );
  };

  if (sessionInitializing) {
    return (
      <div className={styles.chatRoot}>
        <div className={styles.spinnerOverlay}>
          <div className={styles.spinner} />
          {/* <div className={styles.spinnerText}>Initializing Genie session…</div> */}
        </div>
      </div>
    );
  }

  return (
    <div
      className={styles.chatRoot}
      role="region"
      aria-label="Chat conversation"
    >
      <div
        className={styles.messageListWrapper}
        ref={listRef}
        aria-live="polite"
      >
        {messages.length === 0 && !loading && (
          <div className={styles.emptyState}>
            Ask a question to get started. Your SharePoint Copilot-style
            assistant is here to help.
          </div>
        )}
        {messages.map(renderMessage)}
        {loading && showTyping && (
          <div className={styles.typingRow}>
            <div
              className={`${styles.avatar} ${styles.assistantAvatar}`}
              aria-hidden="true"
            >
              Ge
            </div>
            <div className={styles.bubble}>
              <div
                className={styles.typingDots}
                aria-label="Assistant is typing"
              >
                <span></span>
                <span></span>
                <span></span>
              </div>
            </div>
          </div>
        )}
      </div>
      <div className={styles.inputAreaWrapper}>
        <div className={styles.inputInner}>
          {attachments.length > 0 && (
            <div
              className={styles.attachmentsRow}
              aria-label="Selected attachments"
            >
              {attachments.map((a) => (
                <div
                  key={a.id}
                  className={styles.attachmentChip}
                  title={`${a.name} (${Math.round(a.size / 1024)} KB)`}
                >
                  <span>{a.name}</span>
                  <button
                    type="button"
                    className={styles.removeAttachmentBtn}
                    aria-label={`Remove ${a.name}`}
                    onClick={() => removeAttachment(a.id)}
                  >
                    &times;
                  </button>
                </div>
              ))}
            </div>
          )}
          <textarea
            className={styles.textArea}
            placeholder={
              loading
                ? "Generating response…"
                : "Ask something about your site..."
            }
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={handleKeyDown}
            disabled={loading}
            rows={1}
            aria-label="Chat input"
            ref={textAreaRef}
          />
          {/* <div className={styles.metaBar}>
            <span>Enter to send</span>
            <span className={styles.separator}></span>
            <span>Shift+Enter for newline</span>
          </div> */}
          <div className={styles.sendBar}>
            <input
              ref={fileInputRef}
              type="file"
              style={{ display: "none" }}
              multiple
              onChange={(e) => onSelectFiles(e.target.files)}
              aria-hidden="true"
            />
            <button
              type="button"
              className={styles.attachButton}
              onClick={() => fileInputRef.current?.click()}
              disabled={loading}
              aria-label="Attach files"
            >
              <span className={styles.attachIcon} aria-hidden="true">
                <svg
                  viewBox="0 0 24 24"
                  width="16"
                  height="16"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                >
                  <path d="M21.44 11.05l-9.19 9.19a5.5 5.5 0 01-7.78-7.78l9.19-9.19a3.5 3.5 0 014.95 4.95l-9.2 9.19a1.5 1.5 0 01-2.12-2.12l8.49-8.48" />
                </svg>
              </span>
            </button>
            <button
              type="button"
              className={styles.sendButton}
              onClick={handleSend}
              disabled={loading || !input.trim()}
              aria-label="Send message"
            >
              Send
            </button>
          </div>
        </div>
        {error && (
          <div style={{ color: "crimson", fontSize: 12, marginTop: 8 }}>
            {error}
          </div>
        )}
      </div>
    </div>
  );
};

export default Chat;
