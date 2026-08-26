import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { VitePWA } from "vite-plugin-pwa";

export default defineConfig({
  plugins: [
    react(),
    VitePWA({
      // "autoUpdate": il service worker scarica la nuova versione in background
      // e la attiva da solo al prossimo caricamento, senza chiedere nulla
      // all'utente e senza mai toccare i dati salvati in IndexedDB.
      registerType: "autoUpdate",
      includeAssets: ["favicon.ico", "favicon-16.png", "favicon-32.png", "apple-touch-icon.png"],
      manifest: {
        name: "Football Club Manager",
        short_name: "FC Manager",
        description: "Gestionale per la squadra: giocatori, allenamenti, partite, campionato e dossier.",
        start_url: "/",
        scope: "/",
        display: "standalone",
        background_color: "#020617",
        theme_color: "#020617",
        icons: [
          { src: "/favicon-192.png", sizes: "192x192", type: "image/png", purpose: "any" },
          { src: "/icon-512.png", sizes: "512x512", type: "image/png", purpose: "any" },
          { src: "/icon-512-maskable.png", sizes: "512x512", type: "image/png", purpose: "maskable" },
        ],
      },
      workbox: {
        // Precarica tutti i file dell'app costruiti da Vite (JS/CSS/HTML), così
        // il primo avvio offline funziona anche appena dopo l'installazione.
        globPatterns: ["**/*.{js,css,html,png,ico,svg}"],
      },
    }),
  ],
});
