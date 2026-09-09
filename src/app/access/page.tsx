"use client";

import { FormEvent, useState } from "react";
import { useRouter } from "next/navigation";
import styles from "./page.module.css";

export default function AccessPage() {
  const router = useRouter();
  const [password, setPassword] = useState("");

  async function submit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const response = await fetch("/api/access", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ password }) });
    if (!response.ok) {
      setPassword("");
      return;
    }
    router.push("/");
  }

  return (
    <main className={styles.page}>
      <form className={styles.modal} onSubmit={submit} autoComplete="off">
        <span className={styles.brand}>A365</span>
        <h1>Acceso al dashboard</h1>
        <p>Ingresa la contraseña para continuar.</p>
        <label htmlFor="access-password">Contraseña</label>
        <input id="access-password" type="password" value={password} onChange={(event) => setPassword(event.target.value)} autoFocus />
        <button type="submit">Ingresar</button>
      </form>
    </main>
  );
}
