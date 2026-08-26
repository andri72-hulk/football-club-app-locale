import React, { useState, useEffect, useMemo, useCallback } from "react";
import {
  RadarChart, PolarGrid, PolarAngleAxis, PolarRadiusAxis, Radar,
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, Legend
} from "recharts";
import * as XLSX from "xlsx";
import {
  isFileSyncSupported, loadSyncHandle, createSyncFile, linkExistingSyncFile,
  clearSyncHandle, verifySyncPermission, readSyncFile, writeSyncFile,
} from "./fileSync.js";
import {
  Users, Calendar, Trophy, Settings, LayoutDashboard, Plus, X, Search,
  Filter, ChevronRight, ChevronLeft, Clock, MapPin, Shield, Activity,
  Download, Upload, Cloud, CloudOff, Save, Trash2, Edit2, CheckCircle2,
  XCircle, AlertCircle, HeartPulse, Target, Zap, FileSpreadsheet, FileText,
  ArrowLeftRight, Award, Flag, UserCheck, UserX, Menu, ChevronDown, ChevronUp,
  TrendingUp, RefreshCw, Info, LayoutGrid, Handshake, StickyNote, SlidersHorizontal,
  Eye, Maximize2, Share2, Sun, Moon
} from "lucide-react";

/* ============================================================
   COSTANTI E UTILITY
   ============================================================ */

const ROLES = ["Portiere", "Difensore", "Centrocampista", "Attaccante"];

const ROLE_COLORS = {
  Portiere: "bg-amber-500/15 text-amber-400 border-amber-500/30",
  Difensore: "bg-sky-500/15 text-sky-400 border-sky-500/30",
  Centrocampista: "bg-emerald-500/15 text-emerald-400 border-emerald-500/30",
  Attaccante: "bg-rose-500/15 text-rose-400 border-rose-500/30",
};

const POSITIONS = ["Centro", "Destra", "Sinistra"];
const FEET = ["Destro", "Sinistro", "Ambidestro"];

const MEDICAL_STATUS = ["Disponibile", "In dubbio", "Infortunato", "Squalificato"];

const MEDICAL_COLORS = {
  Disponibile: "bg-emerald-500/15 text-emerald-400 border-emerald-500/30",
  "In dubbio": "bg-amber-500/15 text-amber-400 border-amber-500/30",
  Infortunato: "bg-rose-500/15 text-rose-400 border-rose-500/30",
  Squalificato: "bg-slate-500/15 text-slate-400 border-slate-500/30",
};

const ATTENDANCE_STATUS = ["Presente", "Assente", "Giustificato", "Infortunato"];

const ATTENDANCE_COLORS = {
  Presente: "bg-emerald-500/15 text-emerald-400 border-emerald-500/30",
  Assente: "bg-rose-500/15 text-rose-400 border-rose-500/30",
  Giustificato: "bg-amber-500/15 text-amber-400 border-amber-500/30",
  Infortunato: "bg-slate-500/15 text-slate-400 border-slate-500/30",
};

const EXERCISE_TYPE_STYLES = {
  "Attivazione": "bg-emerald-500/15 text-emerald-400 border-emerald-500/30",
  "Tecnica": "bg-yellow-500/15 text-yellow-400 border-yellow-500/30",
  "Tattica": "bg-sky-500/15 text-sky-400 border-sky-500/30",
  "Coordinazione": "bg-purple-500/15 text-purple-400 border-purple-500/30",
  "Decision-Making": "bg-pink-500/15 text-pink-400 border-pink-500/30",
  "Situazionale": "bg-indigo-500/15 text-indigo-400 border-indigo-500/30",
  "Minigiochi": "bg-lime-500/15 text-lime-400 border-lime-500/30",
  "Palla Inattiva": "bg-orange-500/15 text-orange-400 border-orange-500/30",
  "Partita": "bg-cyan-500/15 text-cyan-400 border-cyan-500/30",
  "Defaticamento": "bg-slate-500/15 text-slate-400 border-slate-500/30",
};
const EXERCISE_TYPES = Object.keys(EXERCISE_TYPE_STYLES);

// Esercizi per la categoria Esordienti (Leva 2014, U12-U13), rielaborati a partire da
// metodologie di allenamento giovanile diffuse online (YouCoach, portali di settore giovanile FIGC).
const WEB_EXERCISES_ESORDIENTI = [
  {
    id: "web-ex-1",
    title: "Corsa e Torsioni",
    type: "Attivazione",
    time: "8 min",
    goal: "Mobilità articolare e attivazione neuromuscolare generale",
    description: "Corsa blanda in area libera intervallata da torsioni del busto, slanci delle braccia, skip e calciata dietro. Nessun pallone: prepara il corpo in modo globale prima del lavoro tecnico.",
  },
  {
    id: "web-ex-2",
    title: "Torello ad Alta Intensità",
    type: "Attivazione",
    time: "8 min",
    goal: "Attivazione con la palla e reattività nel piccolo spazio",
    description: "In un quadrato di circa 5 metri, 4 giocatori mantengono il possesso a 1-2 tocchi contro 1 pressatore centrale. Cambio del pressatore ad ogni palla persa.",
  },
  {
    id: "web-ex-3",
    title: "Stazioni di Passaggio e Controllo",
    type: "Tecnica",
    time: "12 min",
    goal: "Migliorare la trasmissione rasoterra e il primo controllo orientato",
    description: "Quadrato di 15 metri con un giocatore ad ogni vertice: passaggio deciso rasoterra al compagno di fronte, che effettua un controllo orientato verso il vertice successivo prima di ripetere l'azione.",
  },
  {
    id: "web-ex-4",
    title: "Controllo a Volo e Tiro",
    type: "Tecnica",
    time: "15 min",
    goal: "Allenare il controllo di palloni alti e la conclusione a rete",
    description: "A gruppi, davanti a 2-3 mini porte poste a 10-15 metri: il mister lancia o calcia il pallone in modo che arrivi al volo o spiovente, il giocatore lo controlla (a volo o orientato) e conclude verso la porta di riferimento.",
  },
  {
    id: "web-ex-5",
    title: "Percorso Scaletta e Ostacoli",
    type: "Coordinazione",
    time: "10 min",
    goal: "Sviluppare le capacità coordinative generali con e senza palla",
    description: "Circuito con scaletta di agilità, cerchi e piccoli ostacoli: prima ripetizione senza pallone concentrandosi sull'appoggio dei piedi, seconda ripetizione conducendo la palla lungo lo stesso percorso.",
  },
  {
    id: "web-ex-6",
    title: "Palla e Ostacoli in Movimento",
    type: "Coordinazione",
    time: "10 min",
    goal: "Coordinazione oculo-podalica abbinata alla conduzione di palla",
    description: "I giocatori conducono il pallone superando una fila di ostacoli bassi; al segnale del mister eseguono un gesto coordinativo (cambio di direzione, passo incrociato, skip) prima di riprendere la conduzione.",
  },
  {
    id: "web-ex-7",
    title: "Dai e Vai in Coppia",
    type: "Tattica",
    time: "12 min",
    goal: "Collaborazione tra due giocatori e inserimento senza palla",
    description: "In coppia: il giocatore passa al compagno e si inserisce immediatamente in profondità per ricevere il pallone di ritorno. Si alternano i lati del campo per lavorare su entrambi i piedi.",
  },
  {
    id: "web-ex-8",
    title: "Ricezione Orientata e Cambio Gioco",
    type: "Tattica",
    time: "12 min",
    goal: "Sviluppare l'apertura di gioco dopo un controllo orientato",
    description: "In uno spazio ampio, il giocatore riceve dando le spalle alla porta avversaria, si gira con un controllo orientato verso il lato libero e apre il gioco sul compagno posizionato dal lato opposto.",
  },
  {
    id: "web-ex-9",
    title: "Palla Viva 4v4 a Tema",
    type: "Decision-Making",
    time: "18 min",
    goal: "Velocizzare la scelta tra passaggio, conduzione e tiro",
    description: "4 contro 4 con due mini porte per squadra. L'allenatore introduce a voce un vincolo a rotazione (es. massimo due tocchi, solo piede debole) per stimolare decisioni rapide sotto pressione.",
  },
  {
    id: "web-ex-10",
    title: "3 Contro 3 Più Jolly",
    type: "Decision-Making",
    time: "15 min",
    goal: "Gestire la superiorità numerica e scegliere il compagno smarcato",
    description: "Due squadre da 3 giocatori più un jolly che gioca sempre con la squadra in possesso, creando superiorità numerica continua e stimolando la lettura dello spazio libero.",
  },
  {
    id: "web-ex-11",
    title: "2 Contro 2 con Sponde Laterali",
    type: "Situazionale",
    time: "15 min",
    goal: "Collaborazione interna ed esterna in spazio ridotto",
    description: "Campo stretto con due jolly laterali utilizzabili come appoggio da entrambe le squadre; il 2 contro 2 centrale gioca per la conclusione in una delle due mini porte.",
  },
  {
    id: "web-ex-12",
    title: "Partenza dal Basso 4 Contro 2",
    type: "Situazionale",
    time: "15 min",
    goal: "Costruzione dal portiere sotto pressione avversaria",
    description: "Portiere e 4 difensori costruiscono l'azione partendo da dietro contro 2 pressatori avversari, cercando di superare la prima linea di pressione con passaggi rapidi e appoggi corti.",
  },
  {
    id: "web-ex-13",
    title: "Gioco dei Colori",
    type: "Minigiochi",
    time: "10 min",
    goal: "Reattività, orientamento spaziale e controllo palla",
    description: "I giocatori conducono liberamente il proprio pallone in un'area delimitata da cinesini colorati; al richiamo di un colore devono raggiungere il cinesino corrispondente nel minor tempo possibile mantenendo il controllo.",
  },
  {
    id: "web-ex-14",
    title: "Guardia del Tesoro",
    type: "Minigiochi",
    time: "10 min",
    goal: "Dribbling e protezione della palla sotto pressione",
    description: "Ogni giocatore conduce il proprio pallone in un'area comune, cercando di calciare fuori i palloni degli avversari mentre protegge il proprio con il corpo.",
  },
  {
    id: "web-ex-15",
    title: "Schema su Calcio d'Angolo",
    type: "Palla Inattiva",
    time: "10 min",
    goal: "Automatismi offensivi sui calci d'angolo",
    description: "Due giocatori si smarcano in contemporanea sul primo palo partendo affiancati in linea, mentre un terzo attacca lo spazio al centro dell'area per la deviazione.",
  },
  {
    id: "web-ex-16",
    title: "Rimessa Laterale Rapida",
    type: "Palla Inattiva",
    time: "8 min",
    goal: "Velocizzare la ripartenza su rimessa laterale",
    description: "Chi effettua la rimessa cerca un appoggio corto e immediato per ripartire in conduzione, oppure un cambio rapido di lato di gioco se la pressione avversaria è alta.",
  },
  {
    id: "web-ex-17",
    title: "Partitella a Tema 7 Contro 7",
    type: "Partita",
    time: "20 min",
    goal: "Applicare in gara i principi allenati nella seduta",
    description: "Partita libera su campo ridotto con un'unica regola tematica legata al lavoro della seduta (es. costruzione dal basso obbligatoria prima di poter attaccare), per il resto gioco libero.",
  },
  {
    id: "web-ex-18",
    title: "Stretching e Riflessione Collettiva",
    type: "Defaticamento",
    time: "6 min",
    goal: "Recupero attivo e consapevolezza del lavoro svolto",
    description: "Allungamento leggero dei principali gruppi muscolari a coppie, seguito da un breve confronto di gruppo su cosa si è imparato durante l'allenamento.",
  },
];

// Nuovi esercizi per Pulcini (8-10 anni), Esordienti (10-12 anni) e Giovanissimi U13/U14/U15 (12-15 anni),
// rielaborati a partire da metodologie di allenamento giovanile pubblicate online (YouCoach, Allenatore.net e portali simili).
const WEB_EXERCISES_MULTI_CATEGORY = [
  // PULCINI 8-10 anni
  {
    id: "web2-ex-1",
    title: "Staffetta Guida e Tiro",
    type: "Coordinazione",
    time: "10 min",
    goal: "[Pulcini 8-10] Coordinazione generale e collaborazione a squadre",
    description: "Conduzione a slalom tra i coni fino a un cerchio, poi tiro in una porticina. Staffetta a squadre con classifica finale, per stimolare divertimento e cooperazione tipici di questa fascia d'età.",
  },
  {
    id: "web2-ex-2",
    title: "Percezione dell'Avversario",
    type: "Minigiochi",
    time: "10 min",
    goal: "[Pulcini 8-10] Percezione dell'avversario e dei compagni di squadra",
    description: "In piccoli gruppi, un giocatore senza palla deve toccare i compagni in conduzione; chi viene toccato cambia ruolo. Aiuta il passaggio dal pensiero egocentrico ('mio') a quello di squadra ('nostro').",
  },
  {
    id: "web2-ex-3",
    title: "Sovrapposizione Semplificata",
    type: "Tattica",
    time: "12 min",
    goal: "[Pulcini 8-10] Introdurre il concetto di sovrapposizione in forma semplice",
    description: "Situazione 2 contro 1: un compagno si sovrappone esternamente al portatore di palla per ricevere sulla corsia e concludere a rete.",
  },
  {
    id: "web2-ex-4",
    title: "Guida di Palla in Velocità",
    type: "Tecnica",
    time: "10 min",
    goal: "[Pulcini 8-10] Velocità di conduzione e controllo della palla",
    description: "Percorso di conduzione rapida tra coni disposti a corridoio, con cambio di direzione al segnale vocale del mister.",
  },
  {
    id: "web2-ex-5",
    title: "2 Contro 2 Libero",
    type: "Minigiochi",
    time: "12 min",
    goal: "[Pulcini 8-10] Primo approccio al gioco situazionale in piccoli gruppi",
    description: "2 contro 2 in spazio ridotto con due porticine per squadra: totale libertà di soluzione tecnica, senza vincoli imposti dal mister.",
  },
  {
    id: "web2-ex-6",
    title: "Dominio della Palla del Portiere",
    type: "Tecnica",
    time: "10 min",
    goal: "[Pulcini 8-10] Dominio della palla con entrambi i piedi per i portieri",
    description: "Esercizio individuale per portieri: controllo e passaggio contro una parete o con il mister, alternando sistematicamente piede destro e sinistro.",
  },

  // ESORDIENTI 10-12 anni (esercizi aggiuntivi rispetto ai 18 già presenti)
  {
    id: "web2-ex-7",
    title: "Trasmissione e Primo Controllo a Tema",
    type: "Tecnica",
    time: "12 min",
    goal: "[Esordienti 10-12] Qualità del passaggio con uno scopo preciso di gioco",
    description: "In gruppi da 3, il ricevente deve orientare il primo controllo verso la porta avversaria prima di restituire il passaggio: il passaggio diventa funzionale al gioco, non fine a se stesso.",
  },
  {
    id: "web2-ex-8",
    title: "Attacco della Profondità",
    type: "Situazionale",
    time: "15 min",
    goal: "[Esordienti 10-12] Attaccare lo spazio alle spalle della difesa",
    description: "Esercizio a tema con lanci filtranti su movimento degli attaccanti ad attaccare la profondità alle spalle di due difensori.",
  },
  {
    id: "web2-ex-9",
    title: "Gioco e Vado",
    type: "Tecnica",
    time: "10 min",
    goal: "[Esordienti 10-12] Abilità tecnica nel passaggio seguito da inserimento",
    description: "Il giocatore passa la palla e segue immediatamente la propria traiettoria per riproporsi al compagno in appoggio, in stile 'gioco e vado'.",
  },
  {
    id: "web2-ex-10",
    title: "Gioco di Posizione a Numeri Ridotti",
    type: "Decision-Making",
    time: "15 min",
    goal: "[Esordienti 10-12] Mobilità e gestione del possesso in piccoli gruppi",
    description: "Possesso palla 4 contro 4 più jolly con zone di riferimento a terra, per stimolare la mobilità e il riconoscimento degli spazi liberi.",
  },
  {
    id: "web2-ex-11",
    title: "Difesa sulla Rimessa Laterale",
    type: "Palla Inattiva",
    time: "8 min",
    goal: "[Esordienti 10-12] Organizzazione difensiva su rimessa laterale avversaria",
    description: "Marcatura dei giocatori vicini al punto di rimessa, con un giocatore incaricato di chiudere lo spazio centrale per impedire l'appoggio corto.",
  },
  {
    id: "web2-ex-12",
    title: "Circuito Mobilità Articolare",
    type: "Defaticamento",
    time: "8 min",
    goal: "[Esordienti 10-12] Prevenzione infortuni e recupero adattato all'età",
    description: "Stazioni di mobilità articolare a corpo libero da eseguire a fine seduta o dopo la partita, con esercizi calibrati sull'età dei ragazzi.",
  },

  // GIOVANISSIMI U13/U14/U15 12-15 anni
  {
    id: "web2-ex-13",
    title: "Gioco di Posizione a Numeri Alti",
    type: "Decision-Making",
    time: "20 min",
    goal: "[U13-U15 12-15] Comprendere il principio di mobilità a numeri alti",
    description: "Possesso palla con più squadre su campo ampio (es. 6 contro 6 più jolly): ricerca costante della mobilità per liberare linee di passaggio.",
  },
  {
    id: "web2-ex-14",
    title: "Piccola Collaborazione Situazionale",
    type: "Situazionale",
    time: "15 min",
    goal: "[U13-U15 12-15] Attivazione mentale e sviluppo tecnico-tattico individuale",
    description: "Esercitazione a numeri ridotti (2 contro 2 o 3 contro 3) in spazio contenuto, con obiettivo di collaborazione rapida tra compagni vicini.",
  },
  {
    id: "web2-ex-15",
    title: "Profondità in un Sistema a Zona",
    type: "Tattica",
    time: "15 min",
    goal: "[U13-U15 12-15] Creare spazio alle spalle della difesa in un sistema strutturato",
    description: "Movimenti a mezzaluna degli attaccanti per attaccare la profondità durante una fase di possesso organizzato, secondo un sistema di gioco a zona (es. 4-3-3 o 4-2-3-1).",
  },
  {
    id: "web2-ex-16",
    title: "Situazioni Rapide di Finalizzazione",
    type: "Decision-Making",
    time: "15 min",
    goal: "[U13-U15 12-15] Soluzioni rapide e creatività nella fase di conclusione",
    description: "Situazioni a numeri ridotti e di breve durata, con l'obiettivo di finalizzare rapidamente l'azione stimolando decisioni immediate e creatività.",
  },
  {
    id: "web2-ex-17",
    title: "Costruzione dal Basso Strutturata",
    type: "Situazionale",
    time: "18 min",
    goal: "[U13-U15 12-15] Coraggio e qualità tecnica nella costruzione contro pressione organizzata",
    description: "Portiere e difesa costruiscono l'azione contro un pressing avversario a uomo strutturato su tutto il fronte offensivo, ricercando il superamento della prima linea.",
  },
  {
    id: "web2-ex-18",
    title: "Capacità Anaerobica con Portieri",
    type: "Attivazione",
    time: "12 min",
    goal: "[U13-U15 12-15] Condizionamento fisico specifico integrato al gioco",
    description: "Piccola partita a tema ad alta intensità (es. 4 contro 4 con portieri), utile per lo sviluppo della capacità anaerobica lattacida in forma ludica.",
  },
];

const MATCH_TYPES = ["Campionato", "Amichevole", "Torneo"];

// Configurazione di default delle tabelle personalizzabili (Configurazioni in Impostazioni)
function defaultConfig() {
  return {
    roles: [...ROLES],
    positions: [...POSITIONS],
    medicalStatuses: [...MEDICAL_STATUS],
    exerciseTypes: [...EXERCISE_TYPES],
    categories: [...DEFAULT_CATEGORIES],
  };
}

// Libreria globale condivisa da TUTTE le stagioni: esercizi, dossier, moduli
// personalizzati e configurazioni. Non vive più dentro le singole stagioni,
// così sopravvive alla creazione di nuove stagioni e all'azzeramento dati.
function defaultLibrary() {
  return {
    exercises: [],
    dossier: [],
    customFormations: [],
    config: defaultConfig(),
  };
}

const ConfigContext = React.createContext(null);
function useConfig() {
  return React.useContext(ConfigContext) || defaultConfig();
}

const NEUTRAL_BADGE = "bg-white/10 text-slate-300 border-white/20";

const MATCH_TYPE_STYLES = {
  Campionato: "bg-emerald-500/15 text-emerald-400 border-emerald-500/30",
  Amichevole: "bg-sky-500/15 text-sky-400 border-sky-500/30",
  Torneo: "bg-amber-500/15 text-amber-400 border-amber-500/30",
};
const MATCH_TYPE_ICONS = { Campionato: Trophy, Amichevole: Handshake, Torneo: Award };
const MATCH_TYPE_BORDER = {
  Campionato: "#10b981",
  Amichevole: "#0ea5e9",
  Torneo: "#f59e0b",
};

const BASE_STAT_KEYS = [
  { key: "speed", label: "Velocità" },
  { key: "technique", label: "Tecnica" },
  { key: "stamina", label: "Resistenza" },
  { key: "tactics", label: "Tattica" },
  { key: "strength", label: "Forza" },
];

// Statistiche mentali: valide per tutti i ruoli
const MENTAL_STAT_KEYS = [
  { key: "leadership", label: "Leadership" },
  { key: "mentality", label: "Mentalità" },
  { key: "concentration", label: "Concentrazione" },
  { key: "commitment", label: "Impegno" },
  { key: "errorManagement", label: "Gestione errore" },
  { key: "teamSpirit", label: "Spirito di squadra" },
];

// Statistiche tecnico/tattiche: non attive per il ruolo Portiere
const TECH_TACTIC_STAT_KEYS = [
  { key: "coordination", label: "Coordinazione" },
  { key: "ballControl", label: "Controllo/Conduzione di palla" },
  { key: "decisionMaking", label: "Scelta/Visione di Gioco" },
  { key: "dribbling", label: "Dribbling" },
  { key: "offBallMovement", label: "Smarcamento" },
  { key: "passing", label: "Passaggio" },
  { key: "shooting", label: "Tiro" },
  { key: "heading", label: "Colpo di testa" },
];

// Statistiche portiere: attive solo per il ruolo Portiere
const GK_STAT_KEYS = [
  { key: "highCatching", label: "Presa alta" },
  { key: "lowCatching", label: "Presa bassa" },
  { key: "diving", label: "Tuffo" },
  { key: "gkPassing", label: "Passaggio" },
  { key: "distribution", label: "Rinvio" },
  { key: "positioning", label: "Posizionamento" },
  { key: "communication", label: "Comunicazione" },
];

const STAT_MIN = 0;
const STAT_MAX = 10;

// Moduli tattici per il calcio a 9 (fonte: documento fornito dal mister)
const FORMATIONS_9V9 = [
  {
    id: "3-2-3",
    name: "3-2-3",
    subtitle: "Il modulo del coraggio",
    description:
      "Tre difensori, due centrocampisti centrali, tre attaccanti (ala destra, punta centrale, ala sinistra). È il modulo più diffuso nel calcio a 9 giovanile perché garantisce ampiezza in entrambe le fasi.",
    structureLabel: "POR — DC · DC · DC — CM · CM — ALA · PC · ALA",
    note: "Consigliato come punto di partenza per la categoria Esordienti: offre chiarezza di ruoli e si presta bene all'insegnamento della partenza dal basso.",
    strengths: [
      "Ampiezza naturale in attacco grazie alle due ali",
      "Superiorità numerica in mezzo campo con i due CM",
      "Solidità difensiva con tre centrali",
      "Facilita il gioco palla a terra e il possesso",
      "Ideale per squadre tecnicamente dotate",
      "Offre molte linee di passaggio al portatore",
    ],
    weaknesses: [
      "Le ali devono fare doppia fase (difensiva e offensiva)",
      "I CM rischiano di coprire troppo spazio da soli",
      "Vulnerabile al pressing alto sul portiere e sulla difesa",
      "Richiede difensori centrali con buona lettura della posizione",
      "Punta centrale a volte isolata se le ali non supportano",
    ],
    positions: [
      { id: "por", label: "POR", role: "Portiere", x: 50, y: 93 },
      { id: "dc1", label: "DC", role: "Difensore", x: 22, y: 74 },
      { id: "dc2", label: "DC", role: "Difensore", x: 50, y: 78 },
      { id: "dc3", label: "DC", role: "Difensore", x: 78, y: 74 },
      { id: "cm1", label: "CM", role: "Centrocampista", x: 34, y: 52 },
      { id: "cm2", label: "CM", role: "Centrocampista", x: 66, y: 52 },
      { id: "ala1", label: "ALA", role: "Attaccante", x: 13, y: 22 },
      { id: "pc", label: "PC", role: "Attaccante", x: 50, y: 13 },
      { id: "ala2", label: "ALA", role: "Attaccante", x: 87, y: 22 },
    ],
  },
  {
    id: "3-3-2",
    name: "3-3-2",
    subtitle: "L'equilibrio",
    description:
      "Tre difensori, tre centrocampisti (mediano di contenimento + due mezzali o trequartisti), due attaccanti. Garantisce un ottimo equilibrio tra le fasi e consente la superiorità numerica a centrocampo.",
    structureLabel: "POR — DC · DC · DC — MED · MEZ · MEZ — ATT · ATT",
    strengths: [
      "Superiorità numerica in mediana (3 contro 2 avversari)",
      "I due attaccanti creano profondità e copertura reciproca",
      "Mediano di contenimento: schermo davanti alla difesa",
      "Le mezzali possono inserirsi creando la terza punta",
      "Ottimo per squadre con buona fisicità a centrocampo",
    ],
    weaknesses: [
      "Minore ampiezza rispetto al 3-2-3 (no ali pure)",
      "Le fasce sono scoperte: i difensori esterni devono salire",
      "I due attaccanti rischiano di restare isolati",
      "Richiede mezzali con elevata capacità di corsa e inserimento",
      "Più complesso da spiegare e applicare ai ragazzi",
    ],
    positions: [
      { id: "por", label: "POR", role: "Portiere", x: 50, y: 93 },
      { id: "dc1", label: "DC", role: "Difensore", x: 22, y: 74 },
      { id: "dc2", label: "DC", role: "Difensore", x: 50, y: 78 },
      { id: "dc3", label: "DC", role: "Difensore", x: 78, y: 74 },
      { id: "med", label: "MED", role: "Centrocampista", x: 50, y: 58 },
      { id: "mez1", label: "MEZ", role: "Centrocampista", x: 27, y: 42 },
      { id: "mez2", label: "MEZ", role: "Centrocampista", x: 73, y: 42 },
      { id: "att1", label: "ATT", role: "Attaccante", x: 35, y: 14 },
      { id: "att2", label: "ATT", role: "Attaccante", x: 65, y: 14 },
    ],
  },
  {
    id: "4-2-2",
    name: "4-2-2",
    subtitle: "La compattezza",
    description:
      "Quattro difensori (due centrali + due terzini), due centrocampisti centrali, due trequartisti/fantasisti. Modulo più difensivista, ideale quando si affronta un avversario tecnicamente superiore.",
    structureLabel: "POR — TRZ · DC · DC · TRZ — CM · CM — TRQ · TRQ",
    strengths: [
      "Solidità difensiva con quattro uomini dietro",
      "I terzini possono spingere sulle fasce in fase offensiva",
      "Due CM garantiscono filtro e regia davanti alla difesa",
      "Ideale per ripartenze veloci con i due trequartisti",
      "Buona copertura delle corsie laterali",
    ],
    weaknesses: [
      "Un attaccante solo: facilmente controllabile dagli avversari",
      "Scarsa presenza in area avversaria",
      "I trequartisti devono avere qualità tecnica elevata",
      "Rischio di restare troppo bassi e subire pressione",
      "Meno adatto alle squadre abituate al gioco offensivo",
    ],
    positions: [
      { id: "por", label: "POR", role: "Portiere", x: 50, y: 93 },
      { id: "trz1", label: "TRZ", role: "Difensore", x: 13, y: 72 },
      { id: "dc1", label: "DC", role: "Difensore", x: 38, y: 76 },
      { id: "dc2", label: "DC", role: "Difensore", x: 62, y: 76 },
      { id: "trz2", label: "TRZ", role: "Difensore", x: 87, y: 72 },
      { id: "cm1", label: "CM", role: "Centrocampista", x: 35, y: 50 },
      { id: "cm2", label: "CM", role: "Centrocampista", x: 65, y: 50 },
      { id: "trq1", label: "TRQ", role: "Attaccante", x: 32, y: 19 },
      { id: "trq2", label: "TRQ", role: "Attaccante", x: 68, y: 19 },
    ],
  },
  {
    id: "3-1-3-1",
    name: "3-1-3-1",
    subtitle: "L'ibrido dinamico",
    description:
      "Tre difensori, un mediano davanti alla difesa (regista basso), tre centrocampisti offensivi/fantasisti, un centravanti. Il modulo più moderno per il calcio a 9, ispirato alla filosofia posizionale.",
    structureLabel: "POR — DC · DC · DC — REG — TRQ · MEZ · TRQ — CT",
    note: "Per i ragazzi di 12-13 anni, il 3-2-3 rimane la scelta ottimale per la maggior parte delle squadre. Il 3-3-2 è il secondo step ideale. Il 3-1-3-1 è da introdurre progressivamente come evoluzione.",
    strengths: [
      "Regista davanti alla difesa: primo riferimento per la costruzione",
      "I tre trequartisti creano superiorità numerica in zona offensiva",
      "Grande dinamismo: i laterali possono diventare esterni bassi",
      "Ottima copertura delle linee di passaggio in tutte le zone",
      "Stimola la creatività e l'intelligenza tattica dei ragazzi",
    ],
    weaknesses: [
      "Molto complesso per ragazzi di 12-13 anni: richiede tempo",
      "Il regista deve avere visione di gioco superiore agli altri",
      "Il centravanti rischia di essere isolato",
      "Se il mediano viene superato, la difesa è in difficoltà",
      "Necessita di molta fase di allenamento dedicata",
    ],
    positions: [
      { id: "por", label: "POR", role: "Portiere", x: 50, y: 93 },
      { id: "dc1", label: "DC", role: "Difensore", x: 22, y: 74 },
      { id: "dc2", label: "DC", role: "Difensore", x: 50, y: 78 },
      { id: "dc3", label: "DC", role: "Difensore", x: 78, y: 74 },
      { id: "reg", label: "REG", role: "Centrocampista", x: 50, y: 58 },
      { id: "trq1", label: "TRQ", role: "Attaccante", x: 20, y: 34 },
      { id: "mez", label: "MEZ", role: "Centrocampista", x: 50, y: 32 },
      { id: "trq2", label: "TRQ", role: "Attaccante", x: 80, y: 34 },
      { id: "ct", label: "CT", role: "Attaccante", x: 50, y: 11 },
    ],
  },
];

// Moduli tattici per il calcio a 5 (categorie Piccoli Amici / Primi Calci)
const FORMATIONS_5V5 = [
  {
    id: "5v5-1-2-1",
    name: "1-2-1",
    subtitle: "Il rombo",
    description:
      "Un difensore centrale, due centrocampisti laterali e una punta. La disposizione a rombo più diffusa nel calcio a 5, offre equilibrio tra copertura difensiva e presenza offensiva.",
    structureLabel: "POR — DC — CM · CM — PC",
    strengths: [
      "Buon equilibrio tra fase difensiva e offensiva",
      "I due centrocampisti possono alternarsi in copertura e spinta",
      "Facile da comprendere per i più piccoli",
      "La punta ha sempre due riferimenti per il triangolo di passaggio",
    ],
    weaknesses: [
      "Il difensore centrale è solo: deve gestire bene gli uno contro uno",
      "Poca ampiezza se i centrocampisti non allargano la posizione",
      "La punta può restare isolata se il possesso si blocca a centrocampo",
    ],
    positions: [
      { id: "por", label: "POR", role: "Portiere", x: 50, y: 92 },
      { id: "dc", label: "DC", role: "Difensore", x: 50, y: 70 },
      { id: "cm1", label: "CM", role: "Centrocampista", x: 25, y: 46 },
      { id: "cm2", label: "CM", role: "Centrocampista", x: 75, y: 46 },
      { id: "pc", label: "PC", role: "Attaccante", x: 50, y: 16 },
    ],
  },
  {
    id: "5v5-2-2",
    name: "2-2",
    subtitle: "Il quadrato",
    description:
      "Due difensori e due attaccanti disposti su due linee parallele. Modulo semplice e speculare, molto usato nelle categorie di base per la sua chiarezza di ruoli.",
    structureLabel: "POR — DC · DC — ATT · ATT",
    strengths: [
      "Ruoli molto chiari: due dietro, due davanti",
      "Buona copertura laterale in entrambe le fasi",
      "Facilita il gioco a due tocchi e le verticalizzazioni rapide",
      "Semplice da insegnare ai bambini più piccoli",
    ],
    weaknesses: [
      "Nessun giocatore di collegamento tra difesa e attacco",
      "Rischio di due linee troppo lontane tra loro",
      "Difficoltà a costruire l'azione se la prima pressione avversaria è alta",
    ],
    positions: [
      { id: "por", label: "POR", role: "Portiere", x: 50, y: 92 },
      { id: "dc1", label: "DC", role: "Difensore", x: 28, y: 66 },
      { id: "dc2", label: "DC", role: "Difensore", x: 72, y: 66 },
      { id: "att1", label: "ATT", role: "Attaccante", x: 28, y: 26 },
      { id: "att2", label: "ATT", role: "Attaccante", x: 72, y: 26 },
    ],
  },
];

// Moduli tattici per il calcio a 7 (categoria Pulcini)
const FORMATIONS_7V7 = [
  {
    id: "7v7-3-2-1",
    name: "3-2-1",
    subtitle: "La solidità",
    description:
      "Tre difensori, due centrocampisti, un attaccante. È l'equivalente semplificato del 4-4-2 nel calcio a 11: pragmatico e ordinato, molto diffuso nella categoria Pulcini.",
    structureLabel: "POR — DC · DC · DC — CM · CM — ATT",
    strengths: [
      "Grande solidità difensiva con tre centrali",
      "Facile da occupare bene il campo senza inventare troppo",
      "I due centrocampisti si aiutano a vicenda in entrambe le fasi",
      "Chiaro riferimento offensivo sulla punta centrale",
    ],
    weaknesses: [
      "Mancanza di ampiezza naturale: nessun giocatore è largo di ruolo",
      "I centrocampisti devono allargarsi manualmente per occupare le fasce",
      "La punta può restare isolata senza un buon supporto dei centrocampisti",
      "Passaggio più complesso verso i moduli a 9 o 11, mancando una struttura di transizione",
    ],
    positions: [
      { id: "por", label: "POR", role: "Portiere", x: 50, y: 92 },
      { id: "dc1", label: "DC", role: "Difensore", x: 22, y: 70 },
      { id: "dc2", label: "DC", role: "Difensore", x: 50, y: 74 },
      { id: "dc3", label: "DC", role: "Difensore", x: 78, y: 70 },
      { id: "cm1", label: "CM", role: "Centrocampista", x: 34, y: 44 },
      { id: "cm2", label: "CM", role: "Centrocampista", x: 66, y: 44 },
      { id: "att", label: "ATT", role: "Attaccante", x: 50, y: 15 },
    ],
  },
  {
    id: "7v7-2-3-1",
    name: "2-3-1",
    subtitle: "L'equivalente del 4-4-2",
    description:
      "Due difensori, tre centrocampisti (due esterni più un centrale), un attaccante. Considerato da molti tecnici il modulo più completo per il calcio a 7 grazie all'ampiezza naturale.",
    structureLabel: "POR — DC · DC — CE · CC · CE — ATT",
    strengths: [
      "Ampiezza naturale grazie ai due centrocampisti esterni",
      "Centrocampo a tre che garantisce equilibrio tra costruzione e copertura",
      "Buona superiorità numerica a centrocampo contro moduli senza fasce",
      "Transizione più naturale verso i moduli a 9 e 11 giocatori",
    ],
    weaknesses: [
      "Solo due difensori: rischio in campo aperto contro ripartenze veloci",
      "Gli esterni devono coprire tutta la fascia in entrambe le fasi",
      "Un solo attaccante può essere preso in marcatura più facilmente",
    ],
    positions: [
      { id: "por", label: "POR", role: "Portiere", x: 50, y: 92 },
      { id: "dc1", label: "DC", role: "Difensore", x: 32, y: 72 },
      { id: "dc2", label: "DC", role: "Difensore", x: 68, y: 72 },
      { id: "ce1", label: "CE", role: "Centrocampista", x: 14, y: 44 },
      { id: "cc", label: "CC", role: "Centrocampista", x: 50, y: 48 },
      { id: "ce2", label: "CE", role: "Centrocampista", x: 86, y: 44 },
      { id: "att", label: "ATT", role: "Attaccante", x: 50, y: 15 },
    ],
  },
  {
    id: "7v7-2-2-2",
    name: "2-2-2",
    subtitle: "Il doppio binario",
    description:
      "Due difensori, due centrocampisti centrali, due attaccanti. Modulo simmetrico e offensivo, utile per stimolare le combinazioni corte in ogni reparto.",
    structureLabel: "POR — DC · DC — CM · CM — ATT · ATT",
    strengths: [
      "Due punte creano più opzioni offensive e si scambiano posizione",
      "Coppie simmetriche in ogni reparto facilitano la comprensione dei compiti",
      "Buona presenza di uomini nella zona di rifinitura",
      "Adatto a stimolare il gioco in verticale con appoggi rapidi",
    ],
    weaknesses: [
      "Nessuna ampiezza sulle fasce, gioco tendenzialmente centrale",
      "Solo due difensori esposti in caso di ripartenza avversaria",
      "I centrocampisti devono coprire molto campo in orizzontale",
    ],
    positions: [
      { id: "por", label: "POR", role: "Portiere", x: 50, y: 92 },
      { id: "dc1", label: "DC", role: "Difensore", x: 32, y: 72 },
      { id: "dc2", label: "DC", role: "Difensore", x: 68, y: 72 },
      { id: "cm1", label: "CM", role: "Centrocampista", x: 32, y: 46 },
      { id: "cm2", label: "CM", role: "Centrocampista", x: 68, y: 46 },
      { id: "att1", label: "ATT", role: "Attaccante", x: 32, y: 16 },
      { id: "att2", label: "ATT", role: "Attaccante", x: 68, y: 16 },
    ],
  },
];

// Moduli tattici classici per il calcio a 11
const FORMATIONS_11 = [
  {
    id: "11-4-4-2",
    name: "4-4-2",
    subtitle: "Il classico",
    description:
      "Quattro difensori, quattro centrocampisti in linea, due punte. Il modulo più tradizionale del calcio a 11, equilibrato e di semplice lettura per squadre e avversari.",
    structureLabel: "POR — TZ · DC · DC · TZ — CE · CC · CC · CE — ATT · ATT",
    strengths: [
      "Grande equilibrio tra reparti, facile da organizzare",
      "Le due punte si aiutano a vicenda e possono scambiarsi",
      "Buona copertura delle fasce con i centrocampisti esterni",
      "Ampiezza naturale sfruttabile dalle sovrapposizioni dei terzini",
    ],
    weaknesses: [
      "Il centrocampo a quattro in linea può essere sovrastato da un centrocampo a tre avversario",
      "Le due punte rischiano di isolarsi se il centrocampo non supporta",
      "Richiede grande disciplina tattica nelle diagonali difensive",
    ],
    positions: [
      { id: "por", label: "POR", role: "Portiere", x: 50, y: 94 },
      { id: "tz1", label: "TZ", role: "Difensore", x: 12, y: 76 },
      { id: "dc1", label: "DC", role: "Difensore", x: 37, y: 80 },
      { id: "dc2", label: "DC", role: "Difensore", x: 63, y: 80 },
      { id: "tz2", label: "TZ", role: "Difensore", x: 88, y: 76 },
      { id: "ce1", label: "CE", role: "Centrocampista", x: 12, y: 48 },
      { id: "cc1", label: "CC", role: "Centrocampista", x: 38, y: 52 },
      { id: "cc2", label: "CC", role: "Centrocampista", x: 62, y: 52 },
      { id: "ce2", label: "CE", role: "Centrocampista", x: 88, y: 48 },
      { id: "att1", label: "ATT", role: "Attaccante", x: 37, y: 15 },
      { id: "att2", label: "ATT", role: "Attaccante", x: 63, y: 15 },
    ],
  },
  {
    id: "11-4-3-3",
    name: "4-3-3",
    subtitle: "L'offensivo",
    description:
      "Quattro difensori, centrocampo a tre, tridente offensivo con due ali e un centravanti. Modulo moderno orientato al possesso palla e all'ampiezza offensiva.",
    structureLabel: "POR — TZ · DC · DC · TZ — CC · CC · CC — ALA · CT · ALA",
    strengths: [
      "Grande ampiezza offensiva grazie alle due ali",
      "Centrocampo a tre flessibile tra costruzione e recupero palla",
      "Il centravanti ha sempre due sponde larghe per lo scarico",
      "Favorisce il possesso palla e la costruzione dal basso",
    ],
    weaknesses: [
      "Le ali devono garantire un grande lavoro di doppia fase",
      "Vulnerabile sulle fasce se i terzini avanzano senza copertura",
      "Il centrocampo a tre può soffrire numericamente contro un centrocampo a quattro",
    ],
    positions: [
      { id: "por", label: "POR", role: "Portiere", x: 50, y: 94 },
      { id: "tz1", label: "TZ", role: "Difensore", x: 12, y: 76 },
      { id: "dc1", label: "DC", role: "Difensore", x: 37, y: 80 },
      { id: "dc2", label: "DC", role: "Difensore", x: 63, y: 80 },
      { id: "tz2", label: "TZ", role: "Difensore", x: 88, y: 76 },
      { id: "cc1", label: "CC", role: "Centrocampista", x: 30, y: 54 },
      { id: "cc2", label: "CC", role: "Centrocampista", x: 50, y: 58 },
      { id: "cc3", label: "CC", role: "Centrocampista", x: 70, y: 54 },
      { id: "ala1", label: "ALA", role: "Attaccante", x: 15, y: 20 },
      { id: "ct", label: "CT", role: "Attaccante", x: 50, y: 13 },
      { id: "ala2", label: "ALA", role: "Attaccante", x: 85, y: 20 },
    ],
  },
  {
    id: "11-3-5-2",
    name: "3-5-2",
    subtitle: "Il centrocampo folto",
    description:
      "Tre difensori centrali, centrocampo a cinque con due esterni a tutta fascia, due punte. Garantisce superiorità numerica a centrocampo e grande copertura delle corsie laterali.",
    structureLabel: "POR — DC · DC · DC — QT · CC · CC · CC · QT — ATT · ATT",
    strengths: [
      "Superiorità numerica a centrocampo contro moduli a quattro",
      "I quinti a tutta fascia garantiscono ampiezza sia offensiva che difensiva",
      "Tre centrali danno grande solidità nella marcatura delle punte avversarie",
      "Le due punte si scambiano riferimenti creando difficoltà alla difesa avversaria",
    ],
    weaknesses: [
      "I quinti devono avere una capacità di corsa e resistenza molto elevata",
      "Se un quinto viene saltato, la difesa a tre resta scoperta sulla fascia",
      "Modulo dispendioso dal punto di vista fisico su tutta la partita",
    ],
    positions: [
      { id: "por", label: "POR", role: "Portiere", x: 50, y: 94 },
      { id: "dc1", label: "DC", role: "Difensore", x: 27, y: 78 },
      { id: "dc2", label: "DC", role: "Difensore", x: 50, y: 82 },
      { id: "dc3", label: "DC", role: "Difensore", x: 73, y: 78 },
      { id: "qt1", label: "QT", role: "Difensore", x: 8, y: 50 },
      { id: "cc1", label: "CC", role: "Centrocampista", x: 30, y: 54 },
      { id: "cc2", label: "CC", role: "Centrocampista", x: 50, y: 58 },
      { id: "cc3", label: "CC", role: "Centrocampista", x: 70, y: 54 },
      { id: "qt2", label: "QT", role: "Difensore", x: 92, y: 50 },
      { id: "att1", label: "ATT", role: "Attaccante", x: 38, y: 15 },
      { id: "att2", label: "ATT", role: "Attaccante", x: 62, y: 15 },
    ],
  },
  {
    id: "11-4-2-3-1",
    name: "4-2-3-1",
    subtitle: "Il moderno",
    description:
      "Quattro difensori, due mediani di contenimento, tre trequartisti a supporto di un'unica punta centrale. Uno dei moduli più diffusi nel calcio moderno per equilibrio e imprevedibilità.",
    structureLabel: "POR — TZ · DC · DC · TZ — MED · MED — TRQ · TRQ · TRQ — CT",
    strengths: [
      "Doppio mediano che garantisce grande equilibrio e copertura difensiva",
      "Tre trequartisti creano superiorità numerica nella zona di rifinitura",
      "Grande imprevedibilità offensiva grazie ai continui scambi di posizione",
      "Buona compattezza tra i reparti in fase di non possesso",
    ],
    weaknesses: [
      "L'unica punta può essere facilmente isolata se il supporto non è tempestivo",
      "Richiede trequartisti tecnicamente evoluti e con buona corsa",
      "Se i mediani vengono saltati, la difesa resta esposta a ripartenze centrali",
    ],
    positions: [
      { id: "por", label: "POR", role: "Portiere", x: 50, y: 94 },
      { id: "tz1", label: "TZ", role: "Difensore", x: 12, y: 78 },
      { id: "dc1", label: "DC", role: "Difensore", x: 37, y: 82 },
      { id: "dc2", label: "DC", role: "Difensore", x: 63, y: 82 },
      { id: "tz2", label: "TZ", role: "Difensore", x: 88, y: 78 },
      { id: "med1", label: "MED", role: "Centrocampista", x: 36, y: 60 },
      { id: "med2", label: "MED", role: "Centrocampista", x: 64, y: 60 },
      { id: "trq1", label: "TRQ", role: "Attaccante", x: 18, y: 32 },
      { id: "trq2", label: "TRQ", role: "Attaccante", x: 50, y: 28 },
      { id: "trq3", label: "TRQ", role: "Attaccante", x: 82, y: 32 },
      { id: "ct", label: "CT", role: "Attaccante", x: 50, y: 12 },
    ],
  },
  {
    id: "11-3-4-3",
    name: "3-4-3",
    subtitle: "L'aggressivo",
    description:
      "Tre difensori centrali, centrocampo a quattro con due esterni a tutta fascia, tridente offensivo. Modulo aggressivo che punta a soffocare l'avversario con il pressing alto su tutto il campo.",
    structureLabel: "POR — DC · DC · DC — CE · CC · CC · CE — ALA · CT · ALA",
    strengths: [
      "Grande densità a centrocampo con quattro uomini",
      "Ampiezza garantita sia in fase difensiva che offensiva dagli esterni",
      "Pressing alto favorito dai tre attaccanti sulla costruzione avversaria",
      "I tre centrali offrono solidità nel gioco aereo",
    ],
    weaknesses: [
      "Gli esterni devono coprire l'intera fascia con doppia fase molto dispendiosa",
      "Enorme spazio scoperto dietro i tre difensori se gli esterni salgono insieme",
      "Vulnerabile alle ripartenze rapide sulle corsie laterali",
      "Richiede giocatori estremamente polivalenti e resistenti",
    ],
    positions: [
      { id: "por", label: "POR", role: "Portiere", x: 50, y: 94 },
      { id: "dc1", label: "DC", role: "Difensore", x: 27, y: 78 },
      { id: "dc2", label: "DC", role: "Difensore", x: 50, y: 82 },
      { id: "dc3", label: "DC", role: "Difensore", x: 73, y: 78 },
      { id: "ce1", label: "CE", role: "Centrocampista", x: 10, y: 52 },
      { id: "cc1", label: "CC", role: "Centrocampista", x: 38, y: 56 },
      { id: "cc2", label: "CC", role: "Centrocampista", x: 62, y: 56 },
      { id: "ce2", label: "CE", role: "Centrocampista", x: 90, y: 52 },
      { id: "ala1", label: "ALA", role: "Attaccante", x: 15, y: 20 },
      { id: "ct", label: "CT", role: "Attaccante", x: 50, y: 13 },
      { id: "ala2", label: "ALA", role: "Attaccante", x: 85, y: 20 },
    ],
  },
  {
    id: "11-5-3-2",
    name: "5-3-2",
    subtitle: "Il fortino",
    description:
      "Cinque difensori (tre centrali più due quinti a tutta fascia), centrocampo a tre, due punte. Modulo molto difensivo che punta sulla solidità e su rapide ripartenze.",
    structureLabel: "POR — QT · DC · DC · DC · QT — CC · CC · CC — ATT · ATT",
    strengths: [
      "Massima solidità difensiva con cinque uomini dietro",
      "I quinti garantiscono ampiezza offensiva quando spingono",
      "Ottimo per le ripartenze veloci grazie alle due punte",
      "Difficile da attaccare centralmente per gli avversari",
    ],
    weaknesses: [
      "Modulo molto prudente, poco controllo del gioco offensivo",
      "Può perdere il controllo del centrocampo contro moduli con più uomini in mezzo",
      "I quinti devono avere una resistenza fisica molto elevata",
      "Le due punte rischiano l'isolamento se la squadra resta troppo bassa",
    ],
    positions: [
      { id: "por", label: "POR", role: "Portiere", x: 50, y: 94 },
      { id: "qt1", label: "QT", role: "Difensore", x: 8, y: 66 },
      { id: "dc1", label: "DC", role: "Difensore", x: 30, y: 78 },
      { id: "dc2", label: "DC", role: "Difensore", x: 50, y: 82 },
      { id: "dc3", label: "DC", role: "Difensore", x: 70, y: 78 },
      { id: "qt2", label: "QT", role: "Difensore", x: 92, y: 66 },
      { id: "cc1", label: "CC", role: "Centrocampista", x: 30, y: 50 },
      { id: "cc2", label: "CC", role: "Centrocampista", x: 50, y: 54 },
      { id: "cc3", label: "CC", role: "Centrocampista", x: 70, y: 50 },
      { id: "att1", label: "ATT", role: "Attaccante", x: 38, y: 15 },
      { id: "att2", label: "ATT", role: "Attaccante", x: 62, y: 15 },
    ],
  },
  {
    id: "11-4-5-1",
    name: "4-5-1",
    subtitle: "Il muro a centrocampo",
    description:
      "Quattro difensori, centrocampo foltissimo a cinque (tre centrali più due esterni), un'unica punta. Modulo che privilegia il controllo del gioco e la densità in mezzo al campo.",
    structureLabel: "POR — TZ · DC · DC · TZ — CE · CC · CC · CC · CE — CT",
    strengths: [
      "Enorme densità e controllo della zona centrocampo",
      "Ottima solidità difensiva grazie al numero di uomini in mezzo",
      "Favorisce il possesso palla e la gestione dei ritmi di gara",
      "I terzini possono spingere con maggiore libertà avendo copertura",
    ],
    weaknesses: [
      "La punta centrale resta spesso isolata senza supporto immediato",
      "Poca profondità e imprevedibilità in fase offensiva",
      "Dipende molto dagli inserimenti dei centrocampisti per creare gol",
      "Rischio di un atteggiamento troppo attendista",
    ],
    positions: [
      { id: "por", label: "POR", role: "Portiere", x: 50, y: 94 },
      { id: "tz1", label: "TZ", role: "Difensore", x: 12, y: 78 },
      { id: "dc1", label: "DC", role: "Difensore", x: 37, y: 82 },
      { id: "dc2", label: "DC", role: "Difensore", x: 63, y: 82 },
      { id: "tz2", label: "TZ", role: "Difensore", x: 88, y: 78 },
      { id: "ce1", label: "CE", role: "Centrocampista", x: 10, y: 50 },
      { id: "cc1", label: "CC", role: "Centrocampista", x: 32, y: 54 },
      { id: "cc2", label: "CC", role: "Centrocampista", x: 50, y: 58 },
      { id: "cc3", label: "CC", role: "Centrocampista", x: 68, y: 54 },
      { id: "ce2", label: "CE", role: "Centrocampista", x: 90, y: 50 },
      { id: "ct", label: "CT", role: "Attaccante", x: 50, y: 14 },
    ],
  },
  {
    id: "11-4-2-4",
    name: "4-2-4",
    subtitle: "Il vecchio assalto",
    description:
      "Quattro difensori, due mediani di contenimento, quattro attaccanti (due ali e due punte centrali). Modulo storico d'assalto, oggi raro ma efficace per la massima pressione offensiva.",
    structureLabel: "POR — TZ · DC · DC · TZ — MED · MED — ALA · ATT · ATT · ALA",
    strengths: [
      "Pressione offensiva massima con quattro uomini avanzati",
      "Ampiezza totale grazie alle due ali larghe",
      "Grande imprevedibilità e superiorità numerica sulla difesa avversaria",
      "Ottimo per recuperare un risultato in svantaggio",
    ],
    weaknesses: [
      "Solo due centrocampisti devono coprire da soli l'intero centrocampo",
      "Rischio enorme in caso di perdita palla e ripartenza avversaria",
      "Scarso equilibrio complessivo tra i reparti",
      "Richiede mediani con capacità di corsa fuori dal comune",
    ],
    positions: [
      { id: "por", label: "POR", role: "Portiere", x: 50, y: 94 },
      { id: "tz1", label: "TZ", role: "Difensore", x: 12, y: 78 },
      { id: "dc1", label: "DC", role: "Difensore", x: 37, y: 82 },
      { id: "dc2", label: "DC", role: "Difensore", x: 63, y: 82 },
      { id: "tz2", label: "TZ", role: "Difensore", x: 88, y: 78 },
      { id: "med1", label: "MED", role: "Centrocampista", x: 38, y: 58 },
      { id: "med2", label: "MED", role: "Centrocampista", x: 62, y: 58 },
      { id: "ala1", label: "ALA", role: "Attaccante", x: 12, y: 22 },
      { id: "att1", label: "ATT", role: "Attaccante", x: 38, y: 13 },
      { id: "att2", label: "ATT", role: "Attaccante", x: 62, y: 13 },
      { id: "ala2", label: "ALA", role: "Attaccante", x: 88, y: 22 },
    ],
  },
  {
    id: "11-5-4-1",
    name: "5-4-1",
    subtitle: "Il bunker",
    description:
      "Cinque difensori, quattro centrocampisti, un'unica punta. Il modulo più difensivo tra i classici, pensato per resistere contro avversari nettamente superiori e ripartire in contropiede.",
    structureLabel: "POR — QT · DC · DC · DC · QT — CE · CC · CC · CE — CT",
    strengths: [
      "Massima copertura difensiva possibile con undici giocatori",
      "Ideale contro avversari tecnicamente o fisicamente superiori",
      "I quinti danno comunque ampiezza in entrambe le fasi",
      "Ottima base per le ripartenze rapide in contropiede",
    ],
    weaknesses: [
      "Presenza offensiva quasi nulla, un solo riferimento avanzato",
      "La punta centrale resta pressoché isolata per tutta la gara",
      "Difficoltà a recuperare palla in zona avanzata del campo",
      "Rischio di subire un assedio prolungato se non si riparte in modo efficace",
    ],
    positions: [
      { id: "por", label: "POR", role: "Portiere", x: 50, y: 94 },
      { id: "qt1", label: "QT", role: "Difensore", x: 8, y: 66 },
      { id: "dc1", label: "DC", role: "Difensore", x: 30, y: 78 },
      { id: "dc2", label: "DC", role: "Difensore", x: 50, y: 82 },
      { id: "dc3", label: "DC", role: "Difensore", x: 70, y: 78 },
      { id: "qt2", label: "QT", role: "Difensore", x: 92, y: 66 },
      { id: "ce1", label: "CE", role: "Centrocampista", x: 12, y: 48 },
      { id: "cc1", label: "CC", role: "Centrocampista", x: 38, y: 52 },
      { id: "cc2", label: "CC", role: "Centrocampista", x: 62, y: 52 },
      { id: "ce2", label: "CE", role: "Centrocampista", x: 88, y: 48 },
      { id: "ct", label: "CT", role: "Attaccante", x: 50, y: 14 },
    ],
  },
];

const FORMATIONS_BY_FORMAT = {
  "5v5": FORMATIONS_5V5,
  "7v7": FORMATIONS_7V7,
  "9v9": FORMATIONS_9V9,
  "11": FORMATIONS_11,
};

// Elenco piatto di tutti i moduli, usato per ricerche per id indipendenti dal formato
const FORMATIONS = [...FORMATIONS_5V5, ...FORMATIONS_7V7, ...FORMATIONS_9V9, ...FORMATIONS_11];

// Numero di giocatori (incluso portiere) per ciascun formato di campo
const FORMAT_PLAYER_COUNT = { "5v5": 5, "7v7": 7, "9v9": 9, "11": 11 };

// Elenco predefinito delle posizioni in campo: usato sia per la legenda nel menu
// Moduli, sia come lista di scelta quando si crea un nuovo modulo personalizzato.
const FORMATION_POSITION_LIBRARY = [
  { label: "POR", description: "Portiere", role: "Portiere" },
  { label: "DC", description: "Difensore Centrale", role: "Difensore" },
  { label: "LIB", description: "Libero", role: "Difensore" },
  { label: "TZ", description: "Terzino", role: "Difensore" },
  { label: "TRZ", description: "Terzino", role: "Difensore" },
  { label: "QT", description: "Quinto (esterno a tutta fascia)", role: "Difensore" },
  { label: "MED", description: "Mediano", role: "Centrocampista" },
  { label: "REG", description: "Regista", role: "Centrocampista" },
  { label: "MEZ", description: "Mezzala", role: "Centrocampista" },
  { label: "CC", description: "Centrocampista Centrale", role: "Centrocampista" },
  { label: "CM", description: "Centrocampista", role: "Centrocampista" },
  { label: "CE", description: "Centrocampista Esterno", role: "Centrocampista" },
  { label: "ED", description: "Esterno Destro", role: "Centrocampista" },
  { label: "ES", description: "Esterno Sinistro", role: "Centrocampista" },
  { label: "TRQ", description: "Trequartista", role: "Attaccante" },
  { label: "ALA", description: "Ala", role: "Attaccante" },
  { label: "AD", description: "Ala Destra", role: "Attaccante" },
  { label: "AS", description: "Ala Sinistra", role: "Attaccante" },
  { label: "CT", description: "Centravanti", role: "Attaccante" },
  { label: "PC", description: "Punta Centrale", role: "Attaccante" },
  { label: "SP", description: "Seconda Punta", role: "Attaccante" },
  { label: "ATT", description: "Attaccante", role: "Attaccante" },
];

// Categorie età precaricate (modificabili/eliminabili dalle Configurazioni)
const DEFAULT_CATEGORIES = [
  "Piccoli Amici",
  "Primi Calci",
  "Pulcini",
  "Esordienti",
  "Under 14",
  "Under 15 (Giovanissimi)",
  "Under 16/17 (Allievi)",
  "Under 18/19 (Juniores)",
  "Prima Squadra",
];

// Mappa di migrazione: converte le vecchie categorie (con fascia d'età nel nome)
// nelle nuove versioni brevi, sia nell'elenco configurato sia negli esercizi già taggati.
const CATEGORY_LABEL_MIGRATION = {
  "Piccoli Amici: 5–6 anni": "Piccoli Amici",
  "Primi Calci: 7–8 anni": "Primi Calci",
  "Pulcini: 9–10 anni": "Pulcini",
  "Esordienti: 11–12 anni": "Esordienti",
  "Under 14: Circa 13-14 anni": "Under 14",
  "Under 15 (Giovanissimi): Circa 14-15 anni": "Under 15 (Giovanissimi)",
  "Under 16 / Under 17 (Allievi): Circa 15-17 anni": "Under 16/17 (Allievi)",
  "Under 18 / Under 19 (Juniores): Circa 17-19 anni": "Under 18/19 (Juniores)",
};

function migrateCategoryLabel(label) {
  return CATEGORY_LABEL_MIGRATION[label] || label;
}

// Applica la migrazione delle etichette categoria sia all'elenco configurato
// sia agli esercizi già taggati con la vecchia versione, così l'aggiornamento
// non richiede di ritoccare a mano ogni voce.
function migrateLibraryCategories(lib) {
  if (!lib) return lib;
  const config = lib.config || defaultConfig();
  const categories = (config.categories || []).map(migrateCategoryLabel);
  const exercises = (lib.exercises || []).map((ex) =>
    ex.category ? { ...ex, category: migrateCategoryLabel(ex.category) } : ex
  );
  return { ...lib, config: { ...config, categories }, exercises };
}

// MIGRAZIONE (26/08/2026): in passato si potevano creare esercizi "al volo"
// direttamente dentro un Focus Tecnico. Questo causava esercizi visibili in
// libreria come duplicati "Da: <nome Focus>" invece che come veri Esercizi
// Singoli. Da ora i Focus possono solo referenziare esercizi già presenti in
// libreria: questa funzione promuove una tantum a Esercizio Singolo ogni
// esercizio già presente in un Focus di qualunque stagione che non abbia già
// un corrispondente standalone (stesso titolo/tipo/tempo/descrizione), così
// nessun contenuto già inserito va perso. Idempotente: rieseguirla su dati
// già migrati non crea ulteriori duplicati.
function migrateFocusExercisesToLibrary(allSeasons, lib) {
  if (!lib) return lib;
  const currentExercises = [...(lib.exercises || [])];
  const matchesExisting = (ex) =>
    currentExercises.some(
      (e) =>
        (e.title || "").trim().toLowerCase() === (ex.title || "").trim().toLowerCase() &&
        (e.type || "Tecnica") === (ex.type || "Tecnica") &&
        (e.time || "") === (ex.time || "") &&
        (e.description || "") === (ex.description || "")
    );
  (allSeasons || []).forEach((s) => {
    (s.focusTecnici || []).forEach((ft) => {
      (ft.exercises || []).forEach((ex) => {
        if (!ex.title || !ex.title.trim()) return;
        if (!matchesExisting(ex)) {
          currentExercises.push({ ...ex, id: uid("ex") });
        }
      });
    });
  });
  return { ...lib, exercises: currentExercises };
}

function emptyLineup() {
  return { formationId: null, assignments: {}, bench: [] };
}

function uid(prefix = "id") {
  return `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
}


function avatarUrl(seed) {
  return `https://api.dicebear.com/7.x/avataaars/svg?seed=${encodeURIComponent(seed || "player")}&backgroundType=solid&backgroundColor=1e293b`;
}

function playerAvatar(player) {
  return player?.photo || avatarUrl(player?.avatarSeed || player?.name);
}

// Formatta "Nome Cognome" in "Nome C." per le etichette compatte sul campo
function shortName(fullName) {
  if (!fullName) return "";
  const parts = fullName.trim().split(/\s+/);
  if (parts.length === 1) return parts[0];
  return `${parts[0]} ${parts[1][0]}.`;
}

// Ridimensiona un'immagine caricata dall'utente e la converte in base64,
// per mantenere piccola l'occupazione nello storage persistente.
function resizeImageFile(file, maxSize = 320, quality = 0.82) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Lettura file non riuscita"));
    reader.onload = () => {
      const img = new Image();
      img.onerror = () => reject(new Error("Immagine non valida"));
      img.onload = () => {
        let { width, height } = img;
        if (width > height && width > maxSize) {
          height = Math.round((height * maxSize) / width);
          width = maxSize;
        } else if (height > maxSize) {
          width = Math.round((width * maxSize) / height);
          height = maxSize;
        }
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}

function formatDate(iso) {
  if (!iso) return "-";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleDateString("it-IT", { day: "2-digit", month: "short", year: "numeric" });
}

function formatDateTime(iso) {
  if (!iso) return "-";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  return `${d.toLocaleDateString("it-IT", { day: "2-digit", month: "short", year: "numeric" })} alle ${d.toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit" })}`;
}

// Sottrae un numero di minuti a un orario "HH:MM", per calcolare l'orario di
// ritrovo al campo a partire dall'orario di inizio partita.
function subtractMinutesFromTime(timeStr, minutes) {
  if (!timeStr) return null;
  const [h, m] = timeStr.split(":").map(Number);
  if (isNaN(h) || isNaN(m)) return null;
  let total = h * 60 + m - minutes;
  if (total < 0) total += 24 * 60; // partita dopo mezzanotte: caso limite, non dovrebbe capitare per un club giovanile
  const hh = String(Math.floor(total / 60) % 24).padStart(2, "0");
  const mm = String(total % 60).padStart(2, "0");
  return `${hh}:${mm}`;
}

// Promemoria di backup manuale: dato che la sincronizzazione tra dispositivi
// diversi da questo (es. il cellulare) richiede un Esporta/Importa JSON
// manuale, teniamo traccia in localStorage (non nei dati della stagione, per
// non "sporcare" il backup stesso) di quando è stato fatto l'ultimo export e
// se il promemoria è stato rimandato dall'utente.
const BACKUP_REMINDER_DAYS = 7;
const LS_LAST_EXPORT_KEY = "football_club_last_export_at";
const LS_FIRST_USE_KEY = "football_club_first_use_at";
const LS_BACKUP_SNOOZE_KEY = "football_club_backup_snooze_until";

function getLastExportAt() {
  try {
    return localStorage.getItem(LS_LAST_EXPORT_KEY);
  } catch {
    return null;
  }
}
function markExportDone() {
  try {
    localStorage.setItem(LS_LAST_EXPORT_KEY, new Date().toISOString());
  } catch {}
}
function getOrInitFirstUseAt() {
  try {
    let v = localStorage.getItem(LS_FIRST_USE_KEY);
    if (!v) {
      v = new Date().toISOString();
      localStorage.setItem(LS_FIRST_USE_KEY, v);
    }
    return v;
  } catch {
    return new Date().toISOString();
  }
}
function getBackupSnoozeUntil() {
  try {
    return localStorage.getItem(LS_BACKUP_SNOOZE_KEY);
  } catch {
    return null;
  }
}
function snoozeBackupReminder(days) {
  try {
    localStorage.setItem(LS_BACKUP_SNOOZE_KEY, new Date(Date.now() + days * 24 * 60 * 60 * 1000).toISOString());
  } catch {}
}

function formatDateShort(iso) {
  if (!iso) return "-";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleDateString("it-IT", { day: "2-digit", month: "2-digit" });
}

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

// Riporta entro la scala 0-10 eventuali valori residui da versioni precedenti dell'app
function clampStat(value) {
  const n = value ?? 5;
  return Math.min(STAT_MAX, Math.max(STAT_MIN, n));
}

function emptyBaseStats() {
  const o = {};
  BASE_STAT_KEYS.forEach((s) => (o[s.key] = 5));
  return o;
}

function emptyMentalStats() {
  const o = {};
  MENTAL_STAT_KEYS.forEach((s) => (o[s.key] = 5));
  return o;
}

function emptyTechTacticStats() {
  const o = {};
  TECH_TACTIC_STAT_KEYS.forEach((s) => (o[s.key] = 5));
  return o;
}

function emptyGkStats() {
  const o = {};
  GK_STAT_KEYS.forEach((s) => (o[s.key] = 5));
  return o;
}

function computePlayerStats(playerId, trainings, matches) {
  const t = trainings || [];
  const m = matches || [];
  const presenze = t.filter((tr) => tr.attendance?.[playerId] === "Presente").length;
  const assenze = t.filter((tr) => tr.attendance?.[playerId] === "Assente").length;
  const convocazioni = m.filter((match) => (match.convocati || []).includes(playerId)).length;
  const reti = m.reduce((sum, match) => sum + (match.scorers || []).filter((s) => s.playerId === playerId).reduce((a, s) => a + (Number(s.goals) || 0), 0), 0);
  const assist = m.reduce((sum, match) => sum + (match.assists || []).filter((s) => s.playerId === playerId).reduce((a, s) => a + (Number(s.assists) || 0), 0), 0);
  const ammonizioni = m.filter((match) => (match.yellowCards || []).includes(playerId)).length;
  const espulsioni = m.filter((match) => (match.redCards || []).includes(playerId)).length;
  return { presenze, assenze, convocazioni, reti, assist, ammonizioni, espulsioni };
}

function emptyCoachNotes() {
  const o = {};
  BASE_STAT_KEYS.forEach((s) => (o[s.key] = ""));
  return o;
}

function newSeason(name) {
  return {
    id: uid("season"),
    name: name || "Nuova Stagione",
    clubName: "",
    teamName: "",
    leva: "",
    teamFormat: "",
    colorPrimary: "#10b981",
    colorSecondary: "#0f172a",
    players: [],
    trainings: [],
    matches: [],
    focusTecnici: [],
    lineup: emptyLineup(),
    championship: { fase1: null, fase2: null, faseFinale: null },
  };
}

const DOSSIER_CATEGORIES = ["Regolamento", "Metodologia", "Tattica", "Notizie", "Documenti Società", "Referti", "Altro"];

const CHAMPIONSHIP_PHASE_META = {
  fase1: { label: "Fase 1", period: "Ottobre - Dicembre" },
  fase2: { label: "Fase 2", period: "Gennaio - Marzo" },
  faseFinale: { label: "Fase Finale", period: "" },
};

function emptyChampionshipPhase(season) {
  return {
    teams: [
      {
        id: uid("cteam"),
        name: season.teamName || "Nostra squadra",
        colorPrimary: season.colorPrimary || "#10b981",
        colorSecondary: season.colorSecondary || "#0f172a",
        isUs: true,
      },
    ],
    matches: [],
    closed: false,
  };
}

function computeStandings(teams, matches) {
  const table = (teams || []).map((t) => ({ ...t, played: 0, w: 0, d: 0, l: 0, gf: 0, ga: 0, pts: 0 }));
  (matches || [])
    .filter((m) => m.played && m.homeGoals != null && m.awayGoals != null)
    .forEach((m) => {
      const home = table.find((t) => t.id === m.homeTeamId);
      const away = table.find((t) => t.id === m.awayTeamId);
      if (!home || !away) return;
      home.played += 1;
      away.played += 1;
      home.gf += m.homeGoals;
      home.ga += m.awayGoals;
      away.gf += m.awayGoals;
      away.ga += m.homeGoals;
      if (m.homeGoals > m.awayGoals) {
        home.w += 1;
        home.pts += 3;
        away.l += 1;
      } else if (m.homeGoals < m.awayGoals) {
        away.w += 1;
        away.pts += 3;
        home.l += 1;
      } else {
        home.d += 1;
        away.d += 1;
        home.pts += 1;
        away.pts += 1;
      }
    });
  return table.sort((a, b) => b.pts - a.pts || b.gf - b.ga - (a.gf - a.ga) || b.gf - a.gf);
}

// Palette standard di colori sociali tra cui scegliere
const TEAM_FORMAT_OPTIONS = ["5v5", "7v7", "9v9", "11"];

const TEAM_COLOR_PALETTE = [
  { name: "Verde", hex: "#16a34a" },
  { name: "Verde acqua", hex: "#059669" },
  { name: "Blu", hex: "#2563eb" },
  { name: "Celeste", hex: "#0ea5e9" },
  { name: "Azzurro", hex: "#38bdf8" },
  { name: "Navy", hex: "#1e293b" },
  { name: "Viola", hex: "#7c3aed" },
  { name: "Rosa", hex: "#db2777" },
  { name: "Rosso", hex: "#dc2626" },
  { name: "Bordeaux", hex: "#991b1b" },
  { name: "Arancione", hex: "#f97316" },
  { name: "Giallo", hex: "#eab308" },
  { name: "Oro", hex: "#ca8a04" },
  { name: "Grigio", hex: "#64748b" },
  { name: "Nero", hex: "#0f172a" },
  { name: "Bianco", hex: "#f8fafc" },
];

const STORAGE_KEY = "football-club-state";
const SHARE_FLAG_KEY = "football-club-share-mode";

/* ============================================================
   COMPONENTI GENERICI UI
   ============================================================ */

function Badge({ children, className = "" }) {
  return (
    <span className={`inline-flex items-center gap-1 rounded-full border px-2.5 py-0.5 text-xs font-medium whitespace-nowrap ${className}`}>
      {children}
    </span>
  );
}

// Menu a tendina compatto, con l'aspetto di un badge, per modificare un valore direttamente in griglia
function InlinePlayerSelect({ value, options, onChange, className = "" }) {
  const normalized = options.map((o) => (typeof o === "string" ? { value: o, label: o } : o));
  return (
    <select
      value={value}
      onChange={(e) => {
        e.stopPropagation();
        onChange(e.target.value);
      }}
      onClick={(e) => e.stopPropagation()}
      className={`rounded-full border px-2 py-0.5 text-[11px] font-medium cursor-pointer outline-none ${className}`}
    >
      {normalized.map((o) => (
        <option key={o.value} value={o.value} className="bg-slate-900 text-slate-200">
          {o.label}
        </option>
      ))}
    </select>
  );
}

function Card({ children, className = "" }) {
  return (
    <div className={`rounded-2xl border border-white/10 bg-slate-900/60 backdrop-blur-sm ${className}`}>
      {children}
    </div>
  );
}

function SectionTitle({ eyebrow, title, icon: Icon, action }) {
  return (
    <div className="flex items-center justify-between mb-4 flex-wrap gap-3">
      <div>
        {eyebrow && (
          <p className="text-[11px] font-semibold tracking-widest text-emerald-400 uppercase mb-1">
            {eyebrow}
          </p>
        )}
        <h2 className="text-xl font-bold text-slate-100 flex items-center gap-2">
          {Icon && <Icon className="w-5 h-5 text-emerald-400" />}
          {title}
        </h2>
      </div>
      {action}
    </div>
  );
}

function Button({ children, variant = "primary", className = "", ...props }) {
  const variants = {
    primary: "bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-semibold shadow-lg shadow-emerald-500/20",
    secondary: "bg-slate-800 hover:bg-slate-700 text-slate-100 border border-white/10",
    ghost: "hover:bg-white/5 text-slate-300",
    danger: "bg-rose-500/15 hover:bg-rose-500/25 text-rose-400 border border-rose-500/30",
  };
  return (
    <button
      className={`inline-flex items-center justify-center gap-2 rounded-xl px-4 py-2.5 text-sm transition-colors ${variants[variant]} ${className}`}
      {...props}
    >
      {children}
    </button>
  );
}

function Modal({ open, onClose, title, children, wide = false, size }) {
  if (!open) return null;
  const widthClass = size === "xl" ? "sm:max-w-5xl" : wide ? "sm:max-w-3xl" : "sm:max-w-lg";
  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-start justify-center bg-black/70 backdrop-blur-sm p-0">
      <div
        className={`w-full ${widthClass} overflow-y-auto rounded-t-3xl sm:rounded-b-2xl border border-white/10 bg-slate-900 shadow-2xl`}
        style={{ maxHeight: "min(98vh, calc(100vh - 4px))" }}
      >
        <div className="sticky top-0 z-10 flex items-center justify-between border-b border-white/10 bg-slate-900/95 backdrop-blur px-5 py-4">
          <h3 className="text-lg font-bold text-slate-100">{title}</h3>
          <button onClick={onClose} className="rounded-lg p-1.5 hover:bg-white/10 text-slate-400">
            <X className="w-5 h-5" />
          </button>
        </div>
        <div className="p-5">{children}</div>
      </div>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label className="block mb-4">
      <span className="block text-xs font-semibold text-slate-400 uppercase tracking-wide mb-1.5">{label}</span>
      {children}
    </label>
  );
}

const inputClass =
  "w-full rounded-xl border border-white/10 bg-slate-950/60 px-3.5 py-2.5 text-sm text-slate-100 placeholder-slate-500 outline-none focus:border-emerald-500/60 focus:ring-2 focus:ring-emerald-500/20";

function ColorPair({ primary, secondary, size = 20 }) {
  return (
    <span className="inline-flex rounded-md overflow-hidden border border-white/20 shrink-0" style={{ width: size * 2, height: size }}>
      <span style={{ width: size, height: size, backgroundColor: primary || "#10b981" }} />
      <span style={{ width: size, height: size, backgroundColor: secondary || "#0f172a" }} />
    </span>
  );
}

function ColorSwatchPicker({ value, onChange }) {
  return (
    <div className="flex flex-wrap gap-2">
      {TEAM_COLOR_PALETTE.map((c) => (
        <button
          key={c.hex}
          type="button"
          title={c.name}
          onClick={() => onChange(c.hex)}
          style={{ backgroundColor: c.hex }}
          className={`w-8 h-8 rounded-full border-2 transition-transform ${
            value === c.hex ? "border-white scale-110 shadow-lg" : "border-white/20 hover:scale-105"
          }`}
        />
      ))}
    </div>
  );
}

function Slider({ value, onChange, label }) {
  function clamp(v) {
    if (isNaN(v)) return STAT_MIN;
    return Math.min(STAT_MAX, Math.max(STAT_MIN, v));
  }
  return (
    <div className="mb-3">
      <div className="flex items-center justify-between mb-1 gap-2">
        <span className="text-xs font-medium text-slate-300">{label}</span>
        <input
          type="number"
          min={STAT_MIN}
          max={STAT_MAX}
          value={value}
          onChange={(e) => onChange(clamp(Number(e.target.value)))}
          className="w-14 rounded-lg border border-white/10 bg-slate-950/60 px-1.5 py-1 text-center text-xs font-bold text-emerald-400 outline-none focus:border-emerald-500/60"
        />
      </div>
      <input
        type="range"
        min={STAT_MIN}
        max={STAT_MAX}
        value={value}
        onChange={(e) => onChange(clamp(Number(e.target.value)))}
        className="w-full accent-emerald-500 h-1.5"
      />
    </div>
  );
}

function EmptyState({ icon: Icon, text }) {
  return (
    <div className="flex flex-col items-center justify-center py-14 text-center">
      <Icon className="w-10 h-10 text-slate-700 mb-3" />
      <p className="text-sm text-slate-500">{text}</p>
    </div>
  );
}

// Pulsante rosso di azzeramento mirato, con richiesta di conferma inline,
// riutilizzato in Giocatori, Allenamenti, Partite e Campionato.
function SectionResetButton({ label = "Azzera", confirmText = "Confermi l'azzeramento? L'operazione non è reversibile.", onConfirm }) {
  const [confirm, setConfirm] = useState(false);
  if (confirm) {
    return (
      <div className="flex items-center gap-2 rounded-xl border border-rose-500/30 bg-rose-500/10 px-3 py-2 flex-wrap">
        <span className="text-xs text-rose-300">{confirmText}</span>
        <Button variant="secondary" className="px-2 py-1 text-xs" onClick={() => setConfirm(false)}>Annulla</Button>
        <Button
          variant="danger"
          className="px-2 py-1 text-xs"
          onClick={() => {
            onConfirm();
            setConfirm(false);
          }}
        >
          Conferma
        </Button>
      </div>
    );
  }
  return (
    <Button variant="danger" onClick={() => setConfirm(true)}>
      <Trash2 className="w-4 h-4" /> {label}
    </Button>
  );
}

/* ============================================================
   APP PRINCIPALE
   ============================================================ */

function formatTime(d) {
  if (!d) return null;
  return d.toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
}

// Indicatore in alto a destra: mostra a colpo d'occhio se l'ultimo salvataggio
// è andato a buon fine e a che ora, distinguendo salvataggio locale (IndexedDB)
// e aggiornamento del file di sincronizzazione (se collegato).
function SaveStatusIndicator({ saving, storageHealthy, lastSavedAt, syncHandle, lastFileSyncAt }) {
  let icon, label, colorClass;
  if (saving) {
    icon = <RefreshCw className="w-3.5 h-3.5 animate-spin" />;
    label = "Salvataggio...";
    colorClass = "border-amber-500/30 bg-amber-500/10 text-amber-300";
  } else if (!storageHealthy) {
    icon = <AlertCircle className="w-3.5 h-3.5" />;
    label = "Non salvato";
    colorClass = "border-rose-500/30 bg-rose-500/10 text-rose-300";
  } else {
    icon = <CheckCircle2 className="w-3.5 h-3.5" />;
    label = lastSavedAt ? `Salvato ${formatTime(lastSavedAt)}` : "Salvato";
    colorClass = "border-emerald-500/30 bg-emerald-500/10 text-emerald-300";
  }

  const tooltipLines = [
    lastSavedAt ? `Ultimo salvataggio locale: ${formatTime(lastSavedAt)}` : "Nessun salvataggio locale ancora effettuato",
  ];
  if (syncHandle) {
    tooltipLines.push(
      lastFileSyncAt
        ? `File sincronizzato aggiornato: ${formatTime(lastFileSyncAt)}`
        : "File di sincronizzazione collegato, in attesa del primo aggiornamento"
    );
  }

  return (
    <div
      title={tooltipLines.join("\n")}
      className={`flex items-center gap-1.5 rounded-lg border px-2.5 py-1.5 text-xs font-medium shrink-0 ${colorClass}`}
    >
      {icon}
      <span className="hidden sm:inline">{label}</span>
    </div>
  );
}

// Ricerca globale: un'unica casella per trovare qualunque cosa nell'app —
// giocatori, allenamenti, partite, esercizi, focus tecnici, documenti del
// dossier — senza dover ricordare in quale sezione si trova. Selezionando un
// risultato si naviga alla sezione giusta (per i giocatori, si apre anche
// direttamente la scheda del giocatore).
function GlobalSearchModal({ open, onClose, season, library, onGoTo, onJumpToPlayer }) {
  const [query, setQuery] = useState("");

  useEffect(() => {
    if (!open) setQuery("");
  }, [open]);

  if (!open) return null;

  const q = query.trim().toLowerCase();
  const results = [];

  if (q) {
    (season?.players || []).forEach((p) => {
      if (p.name.toLowerCase().includes(q)) {
        results.push({ label: p.name, sublabel: p.role || "Giocatore", icon: Users, onSelect: () => onJumpToPlayer(p.id) });
      }
    });
    (season?.trainings || []).forEach((t) => {
      if ((t.focus || "").toLowerCase().includes(q) || formatDate(t.date).toLowerCase().includes(q)) {
        results.push({ label: t.focus || "Allenamento", sublabel: `Allenamento · ${formatDate(t.date)}`, icon: Activity, onSelect: () => onGoTo("trainings") });
      }
    });
    (season?.matches || []).forEach((m) => {
      if ((m.opponent || "").toLowerCase().includes(q)) {
        results.push({ label: `vs ${m.opponent}`, sublabel: `Partita · ${formatDate(m.date)}`, icon: Trophy, onSelect: () => onGoTo("matches") });
      }
    });
    (library?.exercises || []).forEach((ex) => {
      if ((ex.title || "").toLowerCase().includes(q)) {
        results.push({ label: ex.title, sublabel: "Esercizio", icon: Target, onSelect: () => onGoTo("trainings") });
      }
    });
    (season?.focusTecnici || []).forEach((ft) => {
      if ((ft.title || "").toLowerCase().includes(q)) {
        results.push({ label: ft.title, sublabel: "Focus Tecnico", icon: Target, onSelect: () => onGoTo("trainings") });
      }
    });
    (library?.dossier || []).forEach((d) => {
      if ((d.title || "").toLowerCase().includes(q) || (d.fileName || "").toLowerCase().includes(q)) {
        results.push({ label: d.title, sublabel: `Dossier · ${d.fileName}`, icon: FileText, onSelect: () => onGoTo("dossier") });
      }
    });
  }

  return (
    <div className="fixed inset-0 z-[110] bg-black/70 backdrop-blur-sm flex items-start justify-center pt-24 px-4" onClick={onClose}>
      <div className="w-full max-w-lg rounded-2xl border border-white/10 bg-slate-900 shadow-2xl overflow-hidden" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center gap-2 px-4 py-3 border-b border-white/10">
          <Search className="w-4 h-4 text-slate-500 shrink-0" />
          <input
            autoFocus
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Cerca giocatori, allenamenti, partite, esercizi, focus, dossier..."
            className="flex-1 bg-transparent outline-none text-sm text-slate-100 placeholder-slate-500"
          />
          <button onClick={onClose} className="rounded-lg p-1 hover:bg-white/10 text-slate-400 shrink-0">
            <X className="w-4 h-4" />
          </button>
        </div>
        <div className="max-h-96 overflow-y-auto">
          {!q ? (
            <p className="text-xs text-slate-500 px-4 py-6 text-center">Inizia a digitare per cercare in tutta l'app.</p>
          ) : results.length === 0 ? (
            <p className="text-xs text-slate-500 px-4 py-6 text-center">Nessun risultato per "{query}".</p>
          ) : (
            results.slice(0, 30).map((r, i) => (
              <button
                key={i}
                onClick={() => {
                  r.onSelect();
                  onClose();
                }}
                className="w-full flex items-center gap-3 px-4 py-2.5 hover:bg-white/5 text-left"
              >
                <r.icon className="w-4 h-4 text-emerald-400 shrink-0" />
                <div className="min-w-0">
                  <p className="text-sm text-slate-100 truncate">{r.label}</p>
                  <p className="text-[11px] text-slate-500 truncate">{r.sublabel}</p>
                </div>
              </button>
            ))
          )}
        </div>
      </div>
    </div>
  );
}

export default function FootballClubApp() {
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [storageHealthy, setStorageHealthy] = useState(true);
  const [sharedMode, setSharedMode] = useState(false);
  const [seasons, setSeasons] = useState([]);
  const [activeSeasonId, setActiveSeasonId] = useState(null);
  const [library, setLibrary] = useState(null);
  const [loadError, setLoadError] = useState(null);
  const [syncHandle, setSyncHandle] = useState(null);
  const [syncFileName, setSyncFileName] = useState(null);
  const [syncNeedsPermission, setSyncNeedsPermission] = useState(false);
  const [lastSavedAt, setLastSavedAt] = useState(null);
  const [lastFileSyncAt, setLastFileSyncAt] = useState(null);
  const [backupReminderVisible, setBackupReminderVisible] = useState(false);

  useEffect(() => {
    if (loading) return;
    const snooze = getBackupSnoozeUntil();
    if (snooze && new Date(snooze) > new Date()) {
      setBackupReminderVisible(false);
      return;
    }
    const reference = getLastExportAt() || getOrInitFirstUseAt();
    const daysSince = (Date.now() - new Date(reference).getTime()) / (1000 * 60 * 60 * 24);
    setBackupReminderVisible(daysSince >= BACKUP_REMINDER_DAYS);
  }, [loading]);
  const [tab, setTab] = useState("dashboard");
  const [playersView, setPlayersView] = useState("grid"); // 'grid' | 'board'
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [globalSearchOpen, setGlobalSearchOpen] = useState(false);
  const [jumpToPlayerId, setJumpToPlayerId] = useState(null);
  const [theme, setTheme] = useState(() => {
    try {
      return localStorage.getItem("football_club_theme") || "dark";
    } catch {
      return "dark";
    }
  });

  useEffect(() => {
    document.documentElement.classList.toggle("theme-light", theme === "light");
    try {
      localStorage.setItem("football_club_theme", theme);
    } catch {}
  }, [theme]);
  const [toast, setToast] = useState(null);

  const showToast = useCallback((msg, type = "success", action = null) => {
    setToast({ msg, type, action });
    setTimeout(() => setToast(null), action ? 6000 : 2600);
  }, []);

  /* ---------- Caricamento iniziale (richiamabile anche da "Riprova") ---------- */
  async function loadInitialData() {
    setLoading(true);
    setLoadError(null);
    try {
      let shared = false;
      try {
        const flag = await window.storage.get(SHARE_FLAG_KEY, false);
        shared = flag?.value === "true";
      } catch (e) {
          shared = false;
        }
        setSharedMode(shared);

        let data = null;
        let readFailed = false;
        let readErrorMessage = "";
        try {
          const res = await window.storage.get(STORAGE_KEY, shared);
          data = res?.value ? JSON.parse(res.value) : null;
        } catch (e) {
          if (e?.code === "NOT_FOUND") {
            // Nessun dato salvato: è un utente/progetto genuinamente nuovo, sicuro procedere.
            data = null;
          } else {
            // Errore di rete/server (es. instabilità Supabase): NON dobbiamo mai
            // interpretarlo come "nessun dato" — altrimenti il salvataggio automatico
            // sovrascriverebbe i dati reali con una stagione vuota. Blocchiamo tutto
            // e mostriamo un errore con possibilità di riprovare.
            readFailed = true;
            readErrorMessage = e?.message || "Impossibile contattare il server. Controlla la connessione e riprova.";
          }
        }

        if (readFailed) {
          setLoadError(readErrorMessage);
          return;
        }

        // Se esiste un file di sincronizzazione collegato, confrontiamo la sua
        // data di salvataggio con quella locale e usiamo il più recente dei due
        // (copre il caso "ho lavorato sull'altro dispositivo, poi ho aperto qui").
        try {
          const handle = await loadSyncHandle();
          if (handle) {
            setSyncHandle(handle);
            setSyncFileName(handle.name);
            const granted = await verifySyncPermission(handle, true).catch(() => false);
            if (granted) {
              setSyncNeedsPermission(false);
              const fileData = await readSyncFile(handle).catch(() => null);
              if (fileData && fileData.seasons) {
                const fileTime = fileData.savedAt ? new Date(fileData.savedAt).getTime() : 0;
                const localTime = data?.savedAt ? new Date(data.savedAt).getTime() : 0;
                if (fileTime > localTime) {
                  data = fileData;
                  setLastFileSyncAt(new Date(fileTime));
                } else if (data && localTime >= fileTime) {
                  // I dati locali sono più recenti (o identici): riallineiamo il file.
                  await writeSyncFile(handle, data).catch(() => {});
                  setLastFileSyncAt(new Date());
                }
              } else if (data) {
                // File vuoto/nuovo: lo inizializziamo con i dati locali correnti.
                await writeSyncFile(handle, data).catch(() => {});
                setLastFileSyncAt(new Date());
              }
            } else {
              setSyncNeedsPermission(true);
            }
          }
        } catch (e) {
          // Nessun file collegato o errore non bloccante: si procede solo con i dati locali.
        }

        if (data && data.seasons && data.seasons.length > 0) {
          setSeasons(data.seasons);
          setActiveSeasonId(data.activeSeasonId || data.seasons[0].id);
          if (data.savedAt) setLastSavedAt(new Date(data.savedAt));

          if (data.library) {
            // Assicura che eventuali voci aggiunte in versioni successive (es. categorie)
            // siano comunque presenti anche su librerie salvate da versioni precedenti.
            const rawLibrary = {
              ...defaultLibrary(),
              ...data.library,
              config: { ...defaultConfig(), ...(data.library.config || {}) },
            };
            setLibrary(migrateLibraryCategories(migrateFocusExercisesToLibrary(data.seasons, rawLibrary)));
          } else {
            // MIGRAZIONE: prima apertura dopo l'introduzione della libreria globale.
            // Unisce esercizi/dossier già presenti in ciascuna stagione (deduplicando
            // per id) e recupera la configurazione dalla prima stagione disponibile,
            // così nessun dato già inserito viene perso.
            const mergedExercises = [];
            const mergedDossier = [];
            const seenExerciseIds = new Set();
            const seenDossierIds = new Set();
            data.seasons.forEach((s) => {
              (s.exercises || []).forEach((ex) => {
                if (!seenExerciseIds.has(ex.id)) {
                  seenExerciseIds.add(ex.id);
                  mergedExercises.push(ex);
                }
              });
              (s.dossier || []).forEach((d) => {
                if (!seenDossierIds.has(d.id)) {
                  seenDossierIds.add(d.id);
                  mergedDossier.push(d);
                }
              });
            });
            const configSource = data.seasons.find((s) => s.config)?.config;
            const rawLibrary = {
              exercises: mergedExercises,
              dossier: mergedDossier,
              customFormations: [],
              config: { ...defaultConfig(), ...(configSource || {}) },
            };
            setLibrary(migrateLibraryCategories(migrateFocusExercisesToLibrary(data.seasons, rawLibrary)));
          }
        } else {
          const s = newSeason("Stagione 2026-27");
          setSeasons([s]);
          setActiveSeasonId(s.id);
          setLibrary(defaultLibrary());
        }
      } catch (e) {
        // Errore imprevisto (es. dati corrotti): per sicurezza NON creiamo mai
        // una stagione vuota qui, per non rischiare di sovrascrivere dati reali
        // con il salvataggio automatico. Mostriamo l'errore con "Riprova".
        setLoadError(e?.message || "Si è verificato un errore imprevisto durante il caricamento.");
        return;
      } finally {
        setLoading(false);
      }
  }

  useEffect(() => {
    loadInitialData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /* ---------- Salvataggio automatico (con retry) ---------- */
  const saveTimeoutRef = React.useRef(null);

  const persist = useCallback(
    async (nextSeasons, nextActiveId, shared, nextLibrary) => {
      if (typeof window === "undefined" || !window.storage) {
        showToast("Storage non disponibile in questo ambiente", "error");
        return;
      }
      setSaving(true);
      const payload = JSON.stringify({
        seasons: nextSeasons,
        activeSeasonId: nextActiveId,
        library: nextLibrary,
        savedAt: new Date().toISOString(),
      });
      const targetShared = shared ?? sharedMode;
      const maxAttempts = 3;
      let lastError = null;
      for (let attempt = 1; attempt <= maxAttempts; attempt++) {
        try {
          await window.storage.set(STORAGE_KEY, payload, targetShared);
          lastError = null;
          break;
        } catch (e) {
          lastError = e;
          // breve attesa prima di ritentare (backoff crescente)
          await new Promise((r) => setTimeout(r, 400 * attempt));
        }
      }
      if (lastError) {
        setStorageHealthy((wasHealthy) => {
          if (wasHealthy) {
            const detail = lastError?.message ? `: ${lastError.message}` : "";
            showToast(`Salvataggio automatico non disponibile${detail}. I tuoi dati restano visibili in questa sessione: usa "Esporta Dati" per un backup manuale.`, "error");
          }
          return false;
        });
      } else {
        setStorageHealthy((wasHealthy) => {
          if (!wasHealthy) showToast("Salvataggio automatico ripristinato");
          return true;
        });
        setLastSavedAt(new Date());
        // Se un file di sincronizzazione è collegato, lo aggiorniamo in parallelo
        // (best-effort: un eventuale errore qui non blocca il salvataggio principale).
        if (syncHandle) {
          writeSyncFile(syncHandle, JSON.parse(payload))
            .then(() => setLastFileSyncAt(new Date()))
            .catch(() => {
              setSyncNeedsPermission(true);
            });
        }
      }
      setSaving(false);
    },
    [sharedMode, showToast, syncHandle]
  );

  /* ---------- Sincronizzazione file (Google Drive/OneDrive) ---------- */
  async function handleCreateSyncFile() {
    try {
      const handle = await createSyncFile();
      const payload = { seasons, activeSeasonId, library, savedAt: new Date().toISOString() };
      await writeSyncFile(handle, payload);
      setSyncHandle(handle);
      setSyncFileName(handle.name);
      setSyncNeedsPermission(false);
      setLastFileSyncAt(new Date());
      showToast("File di sincronizzazione creato e collegato");
    } catch (e) {
      if (e?.name !== "AbortError") showToast("Impossibile creare il file di sincronizzazione", "error");
    }
  }

  async function handleLinkExistingSyncFile() {
    try {
      const handle = await linkExistingSyncFile();
      const fileData = await readSyncFile(handle).catch(() => null);
      if (fileData && fileData.seasons) {
        setSeasons(fileData.seasons);
        setActiveSeasonId(fileData.activeSeasonId);
        setLibrary(fileData.library || defaultLibrary());
      }
      setSyncHandle(handle);
      setSyncFileName(handle.name);
      setSyncNeedsPermission(false);
      showToast("File di sincronizzazione collegato");
    } catch (e) {
      if (e?.name !== "AbortError") showToast("Impossibile collegare il file", "error");
    }
  }

  async function handleUnlinkSync() {
    await clearSyncHandle();
    setSyncHandle(null);
    setSyncFileName(null);
    setSyncNeedsPermission(false);
    showToast("File di sincronizzazione scollegato");
  }

  async function handleReconfirmSyncPermission() {
    if (!syncHandle) return;
    const granted = await verifySyncPermission(syncHandle, true).catch(() => false);
    setSyncNeedsPermission(!granted);
    if (granted) showToast("Accesso al file confermato");
  }

  // Riconferma automatica del permesso: il browser concede il permesso di
  // scrittura sul file collegato solo in risposta a un'interazione reale
  // dell'utente (non può essere richiesto in automatico al caricamento).
  // Appena serve una riconferma, agganciamo un ascoltatore "una tantum" su
  // tutta l'app: al primo click ovunque, si tenta di riottenere il permesso,
  // così l'utente non deve andare apposta in "Esporta Dati".
  useEffect(() => {
    if (!syncNeedsPermission || !syncHandle) return;
    const handleFirstClick = () => {
      handleReconfirmSyncPermission();
    };
    document.addEventListener("click", handleFirstClick, { once: true, capture: true });
    return () => {
      document.removeEventListener("click", handleFirstClick, { capture: true });
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [syncNeedsPermission, syncHandle]);

  // Debounce: aspetta che l'utente finisca di modificare prima di salvare,
  // evitando chiamate multiple ravvicinate (es. slider trascinati, digitazione rapida)
  useEffect(() => {
    if (loading || loadError) return;
    if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current);
    saveTimeoutRef.current = setTimeout(() => {
      persist(seasons, activeSeasonId, undefined, library);
    }, 500);
    return () => {
      if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [seasons, activeSeasonId, library, loading]);

  const activeSeason = useMemo(
    () => seasons.find((s) => s.id === activeSeasonId) || seasons[0],
    [seasons, activeSeasonId]
  );

  function updateActiveSeason(updater) {
    setSeasons((prev) =>
      prev.map((s) => (s.id === activeSeasonId ? { ...s, ...updater(s) } : s))
    );
  }

  function updateLibrary(updater) {
    setLibrary((prev) => ({ ...(prev || defaultLibrary()), ...updater(prev || defaultLibrary()) }));
  }

  async function toggleSharedMode() {
    const next = !sharedMode;
    setSharedMode(next);
    try {
      await window.storage.set(SHARE_FLAG_KEY, next ? "true" : "false", false);
      // Ricarica dati dal nuovo scope
      const res = await window.storage.get(STORAGE_KEY, next).catch(() => null);
      const data = res?.value ? JSON.parse(res.value) : null;
      if (data && data.seasons && data.seasons.length > 0) {
        setSeasons(data.seasons);
        setActiveSeasonId(data.activeSeasonId || data.seasons[0].id);
        setLibrary(migrateLibraryCategories(data.library ? { ...defaultLibrary(), ...data.library } : defaultLibrary()));
        showToast(next ? "Modalità condivisa attivata: dati dello staff caricati" : "Modalità personale attivata");
      } else {
        // Nessun dato condiviso esistente: proponi di copiare i dati attuali
        await persist(seasons, activeSeasonId, next, library);
        showToast(next ? "Modalità condivisa attivata: dati attuali copiati" : "Modalità personale attivata");
      }
    } catch (e) {
      const detail = e?.message ? `: ${e.message}` : "";
      showToast(`Impossibile cambiare modalità di condivisione${detail}`, "error");
    }
  }

  if (loadError) {
    return (
      <div className="min-h-screen bg-slate-950 flex items-center justify-center px-4">
        <div className="max-w-sm text-center">
          <AlertCircle className="w-10 h-10 text-rose-500 mx-auto mb-4" />
          <p className="text-white font-semibold mb-2">Impossibile caricare i dati</p>
          <p className="text-sm text-slate-400 mb-1">{loadError}</p>
          <p className="text-xs text-slate-500 mb-5">
            Per sicurezza non è stato creato nulla di nuovo: i tuoi dati salvati restano intatti. Riprova quando la connessione è di nuovo stabile.
          </p>
          <Button onClick={loadInitialData}>
            <RefreshCw className="w-4 h-4" /> Riprova
          </Button>
        </div>
      </div>
    );
  }

  if (loading || !library) {
    return (
      <div className="min-h-screen bg-slate-950 flex items-center justify-center">
        <div className="flex flex-col items-center gap-3">
          <div className="w-10 h-10 border-2 border-emerald-500 border-t-transparent rounded-full animate-spin" />
          <p className="text-slate-400 text-sm">Caricamento dati...</p>
        </div>
      </div>
    );
  }

  const NAV_ITEMS = [
    { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { id: "calendar", label: "Calendario", icon: Calendar },
    { id: "players", label: "Giocatori", icon: Users },
    { id: "trainings", label: "Allenamenti", icon: Activity },
    { id: "matches", label: "Partite", icon: Trophy },
    { id: "formations", label: "Moduli", icon: LayoutGrid },
    { id: "championship", label: "Campionato", icon: Award },
    { id: "dossier", label: "Dossier", icon: FileText },
    { id: "export", label: "Esporta Dati", icon: FileSpreadsheet },
    { id: "configurations", label: "Configurazioni", icon: SlidersHorizontal },
    { id: "settings", label: "Impostazioni Stagione", icon: Settings },
  ];

  return (
    <ConfigContext.Provider value={library?.config || defaultConfig()}>
    <div className="min-h-screen bg-slate-950 text-slate-100 font-sans">
      {/* HEADER */}
      <header className="sticky top-0 z-40 border-b border-white/10 bg-slate-950/90 backdrop-blur">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 py-3 flex items-center justify-between gap-3">
          <div className="flex items-center gap-3 min-w-0">
            <button
              className="lg:hidden rounded-lg p-2 hover:bg-white/10"
              onClick={() => setMobileMenuOpen((v) => !v)}
            >
              <Menu className="w-5 h-5" />
            </button>
            {activeSeason?.logoUrl ? (
              <div className="w-9 h-9 rounded-xl overflow-hidden shadow-lg shrink-0 border border-white/10 bg-slate-900">
                <img src={activeSeason.logoUrl} alt="Logo squadra" className="w-full h-full object-cover" />
              </div>
            ) : (
              <div
                className="w-9 h-9 rounded-xl flex items-center justify-center shadow-lg shrink-0"
                style={{
                  backgroundImage: `linear-gradient(135deg, ${activeSeason?.colorPrimary || "#10b981"}, ${activeSeason?.colorSecondary || "#0f172a"})`,
                }}
              >
                <Shield className="w-5 h-5 text-white" />
              </div>
            )}
            <div className="min-w-0">
              <h1 className="text-lg sm:text-xl font-bold text-slate-100 truncate">
                {activeSeason?.teamName || "Nome Squadra"}
              </h1>
              <p className="text-sm text-slate-400 truncate">
                {activeSeason?.leva ? `${activeSeason.leva} · ` : ""}{activeSeason?.name}{activeSeason?.teamFormat ? ` (${activeSeason.teamFormat})` : ""}
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <button
              onClick={() => setTheme((t) => (t === "light" ? "dark" : "light"))}
              className="rounded-lg p-2 hover:bg-white/10 text-slate-400"
              title={theme === "light" ? "Passa al tema scuro" : "Passa al tema chiaro"}
            >
              {theme === "light" ? <Moon className="w-4.5 h-4.5" /> : <Sun className="w-4.5 h-4.5" />}
            </button>
            <button
              onClick={() => setGlobalSearchOpen(true)}
              className="rounded-lg p-2 hover:bg-white/10 text-slate-400"
              title="Cerca in tutta l'app"
            >
              <Search className="w-4.5 h-4.5" />
            </button>
            <SeasonSwitcher
              seasons={seasons}
              activeSeasonId={activeSeasonId}
              onSwitch={setActiveSeasonId}
              onCreate={(name) => {
                const s = newSeason(name);
                setSeasons((prev) => [...prev, s]);
                setActiveSeasonId(s.id);
                showToast("Nuova stagione creata");
              }}
            />
            {!storageHealthy && (
              <button
                onClick={() => persist(seasons, activeSeasonId)}
                disabled={saving}
                className="hidden sm:flex items-center gap-1.5 rounded-lg border border-rose-500/30 bg-rose-500/10 hover:bg-rose-500/20 px-2.5 py-1.5 text-xs text-rose-300 font-medium disabled:opacity-60"
                title="Forza un nuovo tentativo di salvataggio"
              >
                <RefreshCw className={`w-3.5 h-3.5 ${saving ? "animate-spin" : ""}`} /> Riprova salvataggio
              </button>
            )}
            <SaveStatusIndicator
              saving={saving}
              storageHealthy={storageHealthy}
              lastSavedAt={lastSavedAt}
              syncHandle={syncHandle}
              lastFileSyncAt={lastFileSyncAt}
            />
            <div
              title={
                !storageHealthy
                  ? "Salvataggio automatico non disponibile: usa il backup manuale"
                  : sharedMode
                  ? "Dati condivisi con lo staff"
                  : "Dati personali"
              }
              className={`hidden sm:flex items-center gap-1.5 rounded-lg border px-2.5 py-1.5 text-xs ${
                !storageHealthy ? "border-rose-500/30 text-rose-400" : "border-white/10 text-slate-400"
              }`}
            >
              {saving ? (
                <RefreshCw className="w-3.5 h-3.5 animate-spin text-emerald-400" />
              ) : !storageHealthy ? (
                <AlertCircle className="w-3.5 h-3.5" />
              ) : sharedMode ? (
                <Cloud className="w-3.5 h-3.5 text-emerald-400" />
              ) : (
                <CloudOff className="w-3.5 h-3.5" />
              )}
              {saving ? "Salvataggio..." : !storageHealthy ? "Non salvato" : sharedMode ? "Condiviso" : "Personale"}
            </div>
          </div>
        </div>

        {/* NAV DESKTOP */}
        <nav className="hidden lg:flex max-w-7xl mx-auto px-6 gap-1 border-t border-white/5">
          {NAV_ITEMS.map((item) => (
            <button
              key={item.id}
              onClick={() => setTab(item.id)}
              className={`flex items-center gap-2 px-4 py-3 text-sm font-medium border-b-2 transition-colors ${
                tab === item.id
                  ? "border-emerald-400 text-emerald-400"
                  : "border-transparent text-slate-400 hover:text-slate-200"
              }`}
            >
              <item.icon className="w-4 h-4" />
              {item.label}
            </button>
          ))}
        </nav>

        {/* NAV MOBILE */}
        {mobileMenuOpen && (
          <nav className="lg:hidden border-t border-white/10 px-4 py-2 flex flex-col">
            {NAV_ITEMS.map((item) => (
              <button
                key={item.id}
                onClick={() => {
                  setTab(item.id);
                  setMobileMenuOpen(false);
                }}
                className={`flex items-center gap-3 px-3 py-3 rounded-xl text-sm font-medium ${
                  tab === item.id ? "bg-emerald-500/10 text-emerald-400" : "text-slate-400"
                }`}
              >
                <item.icon className="w-4 h-4" />
                {item.label}
              </button>
            ))}
          </nav>
        )}
      </header>

      {/* NAV MOBILE BOTTOM (comoda per il campo) */}
      <nav className="lg:hidden fixed bottom-0 inset-x-0 z-40 bg-slate-900/95 backdrop-blur border-t border-white/10 flex">
        {NAV_ITEMS.filter((n) => n.id !== "export" && n.id !== "settings" && n.id !== "configurations" && n.id !== "calendar").map((item) => (
          <button
            key={item.id}
            onClick={() => setTab(item.id)}
            className={`flex-1 min-w-0 flex flex-col items-center justify-center gap-0.5 py-2 px-0.5 text-[7.5px] font-medium leading-tight ${
              tab === item.id ? "text-emerald-400" : "text-slate-500"
            }`}
          >
            <item.icon className="w-4 h-4 shrink-0" />
            <span className="truncate w-full text-center">{item.label}</span>
          </button>
        ))}
      </nav>

      {/* CONTENUTO */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 py-6 pb-24 lg:pb-10">
        {!storageHealthy && (
          <div className="mb-5 rounded-xl border border-rose-500/30 bg-rose-500/10 px-4 py-3 flex items-start gap-2.5 flex-wrap">
            <AlertCircle className="w-4.5 h-4.5 text-rose-400 shrink-0 mt-0.5" />
            <div className="text-xs text-rose-300 flex-1" style={{ minWidth: 200 }}>
              <p className="font-semibold mb-0.5">Salvataggio automatico non disponibile al momento</p>
              <p className="text-rose-300/80">
                I dati restano visibili finché resti su questa pagina, ma potrebbero non essere ancora salvati in modo persistente.
                Vai su <span className="font-medium">Esporta Dati → Esporta JSON</span> per un backup manuale finché il servizio non si ripristina.
              </p>
            </div>
            <button
              onClick={() => persist(seasons, activeSeasonId)}
              disabled={saving}
              className="shrink-0 flex items-center gap-1.5 rounded-lg border border-rose-500/40 bg-rose-500/15 hover:bg-rose-500/25 px-3 py-2 text-xs text-rose-200 font-semibold disabled:opacity-60"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${saving ? "animate-spin" : ""}`} /> Riprova salvataggio
            </button>
          </div>
        )}
        {!activeSeason ? (
          <EmptyState icon={Calendar} text="Nessuna stagione configurata." />
        ) : (
        <>
        {storageHealthy && backupReminderVisible && (
          <div className="mb-5 rounded-xl border border-amber-500/30 bg-amber-500/10 px-4 py-3 flex items-start gap-2.5 flex-wrap">
            <Save className="w-4.5 h-4.5 text-amber-400 shrink-0 mt-0.5" />
            <div className="text-xs text-amber-300 flex-1" style={{ minWidth: 200 }}>
              <p className="font-semibold mb-0.5">Promemoria: fai un backup</p>
              <p className="text-amber-300/80">
                Non fai un backup manuale (Esporta JSON) da un po'. Se lavori anche da un altro dispositivo (es. il cellulare), è il momento di allinearli.
              </p>
            </div>
            <div className="flex gap-1.5 shrink-0">
              <button
                onClick={() => { snoozeBackupReminder(3); setBackupReminderVisible(false); }}
                className="flex items-center rounded-lg border border-amber-500/40 bg-transparent hover:bg-amber-500/10 px-3 py-2 text-xs text-amber-300 font-medium"
              >
                Ricordamelo tra 3 giorni
              </button>
              <button
                onClick={() => setTab("export")}
                className="flex items-center gap-1.5 rounded-lg border border-amber-500/40 bg-amber-500/15 hover:bg-amber-500/25 px-3 py-2 text-xs text-amber-200 font-semibold"
              >
                <Download className="w-3.5 h-3.5" /> Vai a Esporta Dati
              </button>
            </div>
          </div>
        )}
        </>
        )}
        {!activeSeason ? null : tab === "dashboard" ? (
          <Dashboard
            season={activeSeason}
            onGoTo={setTab}
            onGoToPlayersBoard={() => {
              setPlayersView("board");
              setTab("players");
            }}
          />
        ) : tab === "calendar" ? (
          <CalendarSection season={activeSeason} onGoTo={setTab} />
        ) : tab === "players" ? (
          <PlayersSection
            season={activeSeason}
            updateSeason={updateActiveSeason}
            showToast={showToast}
            view={playersView}
            setView={setPlayersView}
            jumpToPlayerId={jumpToPlayerId}
            onJumpHandled={() => setJumpToPlayerId(null)}
          />
        ) : tab === "trainings" ? (
          <TrainingsSection season={activeSeason} updateSeason={updateActiveSeason} library={library} updateLibrary={updateLibrary} showToast={showToast} />
        ) : tab === "matches" ? (
          <MatchesSection season={activeSeason} updateSeason={updateActiveSeason} showToast={showToast} />
        ) : tab === "formations" ? (
          <FormationsSection season={activeSeason} updateSeason={updateActiveSeason} library={library} updateLibrary={updateLibrary} showToast={showToast} />
        ) : tab === "championship" ? (
          <ChampionshipSection season={activeSeason} updateSeason={updateActiveSeason} showToast={showToast} />
        ) : tab === "dossier" ? (
          <DossierSection library={library} updateLibrary={updateLibrary} showToast={showToast} />
        ) : tab === "export" ? (
          <>
            <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-5 mb-5">
              <div className="flex items-center gap-2 mb-3">
                {syncHandle ? <Cloud className="w-4 h-4 text-emerald-500" /> : <CloudOff className="w-4 h-4 text-slate-500" />}
                <h3 className="text-sm font-semibold text-white">Sincronizzazione automatica (Google Drive / OneDrive)</h3>
              </div>
              {!isFileSyncSupported() ? (
                <p className="text-xs text-slate-500">
                  Questa funzione richiede Chrome, Edge o Opera — il browser che stai usando non la supporta. I backup manuali (Esporta/Importa JSON) restano comunque disponibili qui sotto.
                </p>
              ) : syncHandle ? (
                <div>
                  <p className="text-xs text-slate-400 mb-3">
                    Collegato a <span className="text-slate-200 font-medium">{syncFileName}</span>. Ad ogni salvataggio, questo file viene aggiornato automaticamente — se si trova in una cartella sincronizzata da Google Drive o OneDrive, sarà disponibile anche sull'altro dispositivo.
                  </p>
                  {syncNeedsPermission && (
                    <div className="flex items-center gap-2 mb-3 text-xs text-amber-400">
                      <AlertCircle className="w-3.5 h-3.5" /> Serve riconfermare l'accesso al file.
                      <button onClick={handleReconfirmSyncPermission} className="underline hover:text-amber-300">Riconferma</button>
                    </div>
                  )}
                  <Button variant="secondary" onClick={handleUnlinkSync}>Scollega file</Button>
                </div>
              ) : (
                <div>
                  <p className="text-xs text-slate-400 mb-3">
                    Collega un file dentro una cartella sincronizzata da Google Drive Desktop o OneDrive: l'app lo terrà aggiornato automaticamente, rendendo i dati disponibili anche sul tuo altro dispositivo, senza bisogno di account o login.
                  </p>
                  <div className="flex flex-wrap gap-2">
                    <Button onClick={handleCreateSyncFile}>Crea nuovo file (primo dispositivo)</Button>
                    <Button variant="secondary" onClick={handleLinkExistingSyncFile}>Usa file già esistente (secondo dispositivo)</Button>
                  </div>
                </div>
              )}
            </div>
            <ExportSection
              seasons={seasons}
              activeSeason={activeSeason}
              setSeasons={setSeasons}
              setActiveSeasonId={setActiveSeasonId}
              library={library}
              setLibrary={setLibrary}
              showToast={showToast}
              lastSavedAt={lastSavedAt}
              onExported={() => setBackupReminderVisible(false)}
            />
          </>
        ) : tab === "configurations" ? (
          <ConfigurationsSection library={library} updateLibrary={updateLibrary} showToast={showToast} />
        ) : tab === "settings" ? (
          <SettingsSection
            season={activeSeason}
            updateSeason={updateActiveSeason}
            sharedMode={sharedMode}
            onToggleShared={toggleSharedMode}
            seasons={seasons}
            setSeasons={setSeasons}
            activeSeasonId={activeSeasonId}
            setActiveSeasonId={setActiveSeasonId}
            showToast={showToast}
          />
        ) : null}
      </main>

      <GlobalSearchModal
        open={globalSearchOpen}
        onClose={() => setGlobalSearchOpen(false)}
        season={activeSeason}
        library={library}
        onGoTo={setTab}
        onJumpToPlayer={(id) => {
          setJumpToPlayerId(id);
          setTab("players");
        }}
      />

      {toast && (
        <div
          className={`fixed bottom-20 lg:bottom-6 left-1/2 -translate-x-1/2 z-50 rounded-xl px-4 py-3 text-sm font-medium shadow-2xl border flex items-center gap-3 ${
            toast.type === "error"
              ? "bg-rose-950 border-rose-500/40 text-rose-300"
              : "bg-emerald-950 border-emerald-500/40 text-emerald-300"
          }`}
        >
          <span>{toast.msg}</span>
          {toast.action && (
            <button
              onClick={() => {
                toast.action.onClick();
                setToast(null);
              }}
              className="shrink-0 underline font-semibold hover:opacity-80"
            >
              {toast.action.label}
            </button>
          )}
        </div>
      )}
    </div>
    </ConfigContext.Provider>
  );
}

/* ============================================================
   SELETTORE STAGIONE
   ============================================================ */

function SeasonSwitcher({ seasons, activeSeasonId, onSwitch, onCreate }) {
  const [open, setOpen] = useState(false);
  const [creating, setCreating] = useState(false);
  const [name, setName] = useState("");

  return (
    <div className="relative">
      <button
        onClick={() => setOpen((v) => !v)}
        className="flex items-center gap-1.5 rounded-lg border border-white/10 bg-slate-900 px-3 py-2 text-xs font-medium text-slate-300 hover:bg-slate-800"
      >
        <Calendar className="w-3.5 h-3.5 text-emerald-400" />
        <span className="hidden sm:inline max-w-[110px] truncate">
          {seasons.find((s) => s.id === activeSeasonId)?.name || "Stagione"}
        </span>
        <ChevronDown className="w-3.5 h-3.5" />
      </button>
      {open && (
        <div className="absolute right-0 mt-2 w-64 rounded-xl border border-white/10 bg-slate-900 shadow-2xl z-50 overflow-hidden">
          <div className="max-h-56 overflow-y-auto">
            {seasons.map((s) => (
              <button
                key={s.id}
                onClick={() => {
                  onSwitch(s.id);
                  setOpen(false);
                }}
                className={`w-full text-left px-3.5 py-2.5 text-sm hover:bg-white/5 flex items-center justify-between ${
                  s.id === activeSeasonId ? "text-emerald-400" : "text-slate-300"
                }`}
              >
                {s.name}
                {s.id === activeSeasonId && <CheckCircle2 className="w-4 h-4" />}
              </button>
            ))}
          </div>
          <div className="border-t border-white/10 p-2">
            {creating ? (
              <div className="flex gap-1.5">
                <input
                  autoFocus
                  value={name}
                  onChange={(e) => setName(e.target.value)}
                  placeholder="Es. Stagione 2027-28"
                  className={inputClass + " py-1.5"}
                />
                <button
                  onClick={() => {
                    if (name.trim()) {
                      onCreate(name.trim());
                      setName("");
                      setCreating(false);
                      setOpen(false);
                    }
                  }}
                  className="shrink-0 rounded-lg bg-emerald-500 px-2.5 text-slate-950"
                >
                  <Plus className="w-4 h-4" />
                </button>
              </div>
            ) : (
              <button
                onClick={() => setCreating(true)}
                className="w-full flex items-center gap-2 px-2 py-2 text-sm text-emerald-400 hover:bg-white/5 rounded-lg"
              >
                <Plus className="w-4 h-4" /> Nuova stagione
              </button>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

/* ============================================================
   DASHBOARD
   ============================================================ */

function Dashboard({ season, onGoTo, onGoToPlayersBoard }) {
  const players = season.players || [];
  const trainings = [...(season.trainings || [])].sort((a, b) => new Date(a.date) - new Date(b.date));
  const matches = [...(season.matches || [])].sort((a, b) => new Date(a.date) - new Date(b.date));

  const now = new Date();
  const nextTraining = trainings.find((t) => new Date(t.date) >= now) || trainings[trainings.length - 1];
  const nextMatch = matches.find((m) => m.status === "Programmata" && new Date(m.date) >= now);

  const avgAttendance = useMemo(() => {
    const past = trainings.filter((t) => t.attendance && Object.keys(t.attendance).length > 0);
    if (past.length === 0) return 0;
    const totalPct = past.reduce((acc, t) => {
      const values = Object.values(t.attendance);
      const present = values.filter((v) => v === "Presente").length;
      return acc + (values.length ? (present / values.length) * 100 : 0);
    }, 0);
    return Math.round(totalPct / past.length);
  }, [trainings]);

  const record = useMemo(() => {
    const played = matches.filter((m) => m.status === "Disputata" && m.result);
    let v = 0, n = 0, p = 0;
    played.forEach((m) => {
      if (m.result.golFor > m.result.golAgainst) v++;
      else if (m.result.golFor === m.result.golAgainst) n++;
      else p++;
    });
    return { v, n, p, totalPlayed: played.length };
  }, [matches]);

  const recordData = [
    { name: "Vittorie", value: record.v, color: "#10b981" },
    { name: "Pareggi", value: record.n, color: "#f59e0b" },
    { name: "Sconfitte", value: record.p, color: "#f43f5e" },
  ];

  const attendanceTrend = trainings
    .filter((t) => t.attendance && Object.keys(t.attendance).length > 0)
    .slice(-8)
    .map((t) => {
      const values = Object.values(t.attendance);
      const present = values.filter((v) => v === "Presente").length;
      return {
        date: formatDateShort(t.date),
        presenza: values.length ? Math.round((present / values.length) * 100) : 0,
      };
    });

  const upcomingEvents = [
    ...trainings.filter((t) => new Date(t.date) >= now).map((t) => ({ type: "training", date: t.date, data: t })),
    ...matches.filter((m) => m.status === "Programmata" && new Date(m.date) >= now).map((m) => ({ type: "match", date: m.date, data: m })),
  ]
    .sort((a, b) => new Date(a.date) - new Date(b.date))
    .slice(0, 5);

  const activeFormation = FORMATIONS.find((f) => f.id === season.lineup?.formationId);
  const nextTrainingFocus = nextTraining ? (season.focusTecnici || []).find((f) => f.id === nextTraining.focusTecnicoId) : null;

  // Capocannoniere e miglior assistman della squadra
  const { topScorer, topAssist } = useMemo(() => {
    let scorer = null;
    let assist = null;
    players.forEach((p) => {
      const s = computePlayerStats(p.id, [], matches);
      if (s.reti > 0 && (!scorer || s.reti > scorer.value)) scorer = { name: p.name, value: s.reti };
      if (s.assist > 0 && (!assist || s.assist > assist.value)) assist = { name: p.name, value: s.assist };
    });
    return { topScorer: scorer, topAssist: assist };
  }, [players, matches]);

  // Determina la fase di campionato "attiva" per la Dashboard:
  // Fase Finale se creata (niente classifica, solo ultimo risultato), altrimenti Fase 2 se creata, altrimenti Fase 1.
  const championship = season.championship || {};
  let champPhaseKey = null;
  if (championship.faseFinale) champPhaseKey = "faseFinale";
  else if (championship.fase2) champPhaseKey = "fase2";
  else if (championship.fase1) champPhaseKey = "fase1";

  const champBox = useMemo(() => {
    if (!champPhaseKey) return null;
    const phase = championship[champPhaseKey];
    const meta = CHAMPIONSHIP_PHASE_META[champPhaseKey];
    if (champPhaseKey === "faseFinale") {
      const played = [...(phase.matches || [])].filter((m) => m.played).sort((a, b) => new Date(b.date) - new Date(a.date));
      const last = played[0];
      if (!last) return { meta, noData: true };
      const home = phase.teams.find((t) => t.id === last.homeTeamId);
      const away = phase.teams.find((t) => t.id === last.awayTeamId);
      const usHome = home?.isUs;
      const opponent = usHome ? away : home;
      const usGoals = usHome ? last.homeGoals : last.awayGoals;
      const oppGoals = usHome ? last.awayGoals : last.homeGoals;
      return { meta, lastResult: { usGoals, oppGoals, opponentName: opponent?.name || "?" } };
    }
    const standings = computeStandings(phase.teams, phase.matches);
    const position = standings.findIndex((t) => t.isUs) + 1;
    const us = standings.find((t) => t.isUs);
    if (!position) return { meta, noData: true };
    return { meta, position, points: us?.pts ?? 0 };
  }, [championship, champPhaseKey]);

  return (
    <div className="space-y-6">
      {/* KPI ROW */}
      <div className="grid grid-cols-2 lg:grid-cols-3 gap-4">
        <KpiCard icon={Users} label="Giocatori in rosa" value={players.length} accent="emerald" onClick={() => onGoTo("players")} />
        <KpiCard icon={Activity} label="Allenamenti: Presenza Media" value={`${avgAttendance}%`} accent="sky" onClick={() => onGoTo("trainings")} />
        <KpiCard
          icon={Trophy}
          label="Partite disputate"
          value={record.totalPlayed}
          subtext={`Record V-N-P: ${record.v}-${record.n}-${record.p}`}
          subtextClassName="text-sky-400 font-bold"
          accent="amber"
          onClick={() => onGoTo("matches")}
        />
        <KpiCard
          icon={LayoutGrid}
          label="Modulo attivo"
          value={activeFormation ? activeFormation.name : "—"}
          subtext={activeFormation ? activeFormation.subtitle : "Nessuno selezionato"}
          accent="rose"
          onClick={() => onGoTo("formations")}
        />
        <KpiCard
          icon={Target}
          label="Capocannoniere ed Assist"
          value={
            !topScorer && !topAssist ? (
              "—"
            ) : (
              <>
                <div>{topScorer ? `${topScorer.value} Ret${topScorer.value === 1 ? "e" : "i"} (${topScorer.name})` : "—"}</div>
                <div>{topAssist ? `${topAssist.value} Assist (${topAssist.name})` : ""}</div>
              </>
            )
          }
          valueClassName="text-lg sm:text-xl font-extrabold text-slate-100 leading-tight space-y-0.5"
          accent="emerald"
          onClick={onGoToPlayersBoard}
        />
        <KpiCard
          icon={Award}
          label={champBox ? `Campionato — ${champBox.meta.label}` : "Campionato"}
          value={
            !champBox
              ? "—"
              : champBox.noData
              ? "—"
              : champPhaseKey === "faseFinale"
              ? `${champBox.lastResult.usGoals} - ${champBox.lastResult.oppGoals}`
              : `${champBox.position}°`
          }
          subtext={
            !champBox
              ? "Nessuna fase creata"
              : champBox.noData
              ? champPhaseKey === "faseFinale"
                ? "Nessun risultato ancora"
                : "Nessuna partita disputata"
              : champPhaseKey === "faseFinale"
              ? `vs ${champBox.lastResult.opponentName}`
              : `${champBox.points} punti in classifica`
          }
          accent="sky"
          onClick={() => onGoTo("championship")}
        />
      </div>

      {/* NEXT EVENTS */}
      <div className="grid lg:grid-cols-2 gap-4">
        <Card className="p-5">
          <div className="flex items-center gap-2 mb-3">
            <div className="w-9 h-9 rounded-xl bg-sky-500/15 flex items-center justify-center">
              <Activity className="w-4.5 h-4.5 text-sky-400" />
            </div>
            <p className="text-[11px] font-semibold tracking-widest text-sky-400 uppercase">Prossimo Allenamento</p>
          </div>
          {nextTraining ? (
            <div>
              <p className="text-lg font-bold text-slate-100">{formatDate(nextTraining.date)} · {nextTraining.time || "--:--"}</p>
              <p className="text-sm text-slate-400 mt-1">{nextTrainingFocus?.title || nextTraining.focus || "Nessun focus tecnico specificato"}</p>
              {nextTrainingFocus && (nextTrainingFocus.exercises || []).length > 0 && (
                <div className="mt-2 space-y-1">
                  {nextTrainingFocus.exercises.map((ex, i) => (
                    <div key={ex.id || i} className="flex items-center gap-1.5 flex-wrap">
                      <Badge className={EXERCISE_TYPE_STYLES[ex.type] || EXERCISE_TYPE_STYLES.Tecnica}>{ex.type || "Tecnica"}</Badge>
                      <span className="text-xs text-slate-400">{ex.title || `Esercizio ${i + 1}`}</span>
                      <span className="text-xs text-slate-600">· {ex.time || "--"}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          ) : (
            <p className="text-sm text-slate-500">Nessun allenamento programmato.</p>
          )}
        </Card>

        <Card className="p-5">
          <div className="flex items-center gap-2 mb-3">
            <div className="w-9 h-9 rounded-xl bg-emerald-500/15 flex items-center justify-center">
              <Trophy className="w-4.5 h-4.5 text-emerald-400" />
            </div>
            <p className="text-[11px] font-semibold tracking-widest text-emerald-400 uppercase">Prossima Partita</p>
          </div>
          {nextMatch ? (
            <div className="flex items-center justify-between">
              <div>
                <p className="text-lg font-bold text-slate-100">
                  {season.teamName || "Squadra"} <span className="text-slate-500">vs</span> {nextMatch.opponent}
                </p>
                <p className="text-sm text-slate-400 mt-1 flex items-center gap-1.5">
                  <Clock className="w-3.5 h-3.5" /> {formatDate(nextMatch.date)} · {nextMatch.time || "--:--"}
                  <span className="mx-1">·</span>
                  <MapPin className="w-3.5 h-3.5" /> {nextMatch.venue}
                </p>
              </div>
            </div>
          ) : (
            <p className="text-sm text-slate-500">Nessuna partita programmata.</p>
          )}
        </Card>
      </div>

      {/* CHARTS */}
      <div className="grid lg:grid-cols-3 gap-4">
        <Card className="p-5 lg:col-span-2">
          <p className="text-sm font-semibold text-slate-300 mb-3">Andamento presenze (ultimi allenamenti)</p>
          {attendanceTrend.length > 0 ? (
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={attendanceTrend}>
                <CartesianGrid strokeDasharray="3 3" stroke="#ffffff10" />
                <XAxis dataKey="date" stroke="#64748b" fontSize={12} />
                <YAxis stroke="#64748b" fontSize={12} unit="%" />
                <Tooltip contentStyle={{ background: "#0f172a", border: "1px solid #ffffff20", borderRadius: 10 }} />
                <Bar dataKey="presenza" fill="#10b981" radius={[6, 6, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          ) : (
            <EmptyState icon={Activity} text="Nessun dato di presenza disponibile ancora." />
          )}
        </Card>

        <Card className="p-5">
          <p className="text-sm font-semibold text-slate-300 mb-3">Record partite</p>
          {record.totalPlayed > 0 ? (
            <ResponsiveContainer width="100%" height={220}>
              <PieChart>
                <Pie data={recordData} dataKey="value" nameKey="name" innerRadius={45} outerRadius={75} paddingAngle={3}>
                  {recordData.map((d, i) => (
                    <Cell key={i} fill={d.color} />
                  ))}
                </Pie>
                <Legend verticalAlign="bottom" height={24} iconType="circle" wrapperStyle={{ fontSize: 12, color: "#94a3b8" }} />
                <Tooltip contentStyle={{ background: "#0f172a", border: "1px solid #ffffff20", borderRadius: 10 }} />
              </PieChart>
            </ResponsiveContainer>
          ) : (
            <EmptyState icon={Trophy} text="Nessuna partita disputata ancora." />
          )}
        </Card>
      </div>

      {/* UPCOMING EVENTS LIST */}
      <Card className="p-5">
        <p className="text-sm font-semibold text-slate-300 mb-3">Prossimi eventi</p>
        {upcomingEvents.length === 0 ? (
          <EmptyState icon={Calendar} text="Nessun evento in programma." />
        ) : (
          <div className="divide-y divide-white/5">
            {upcomingEvents.map((e, i) => (
              <div key={i} className="flex items-center gap-3 py-3">
                <div
                  className={`w-9 h-9 rounded-xl flex items-center justify-center shrink-0 ${
                    e.type === "training" ? "bg-sky-500/15 text-sky-400" : "bg-emerald-500/15 text-emerald-400"
                  }`}
                >
                  {e.type === "training" ? <Activity className="w-4.5 h-4.5" /> : <Trophy className="w-4.5 h-4.5" />}
                </div>
                <div className="flex-1 min-w-0">
                  <p className="text-sm font-medium text-slate-100 truncate">
                    {e.type === "training" ? e.data.focus || "Allenamento" : `vs ${e.data.opponent}`}
                  </p>
                  <p className="text-xs text-slate-500">{formatDate(e.date)} · {e.data.time || "--:--"}</p>
                </div>
              </div>
            ))}
          </div>
        )}
      </Card>
    </div>
  );
}

function KpiCard({ icon: Icon, label, value, valueClassName, subtext, subtextClassName, accent, onClick }) {
  const accents = {
    emerald: "bg-emerald-500/15 text-emerald-400",
    sky: "bg-sky-500/15 text-sky-400",
    amber: "bg-amber-500/15 text-amber-400",
    rose: "bg-rose-500/15 text-rose-400",
  };
  return (
    <button onClick={onClick} className="text-left">
      <Card className="p-4 hover:border-white/20 transition-colors h-full">
        <div className={`w-9 h-9 rounded-xl flex items-center justify-center mb-3 ${accents[accent]}`}>
          <Icon className="w-4.5 h-4.5" />
        </div>
        <div className={valueClassName || "text-2xl font-extrabold text-slate-100 truncate"}>{value}</div>
        <p className="text-xs text-slate-500 mt-0.5">{label}</p>
        {subtext && <p className={`text-[11px] mt-1 truncate ${subtextClassName || "text-slate-600"}`}>{subtext}</p>}
      </Card>
    </button>
  );
}

/* ============================================================
   SEZIONE CALENDARIO
   ============================================================ */

const ITALIAN_MONTHS = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"];
const ITALIAN_WEEKDAYS = ["Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"];

function sameDay(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

// Vista calendario mensile: allenamenti e partite del mese, con un puntino
// colorato per giorno. Vista di sola consultazione — per aprire il dettaglio
// di un evento si passa comunque dalle sezioni Allenamenti/Partite.
function CalendarSection({ season, onGoTo }) {
  const trainings = season?.trainings || [];
  const matches = season?.matches || [];
  const [cursor, setCursor] = useState(() => new Date());
  const [selectedDay, setSelectedDay] = useState(null);

  const year = cursor.getFullYear();
  const month = cursor.getMonth();

  const firstOfMonth = new Date(year, month, 1);
  const startWeekday = (firstOfMonth.getDay() + 6) % 7; // lunedì = 0
  const daysInMonth = new Date(year, month + 1, 0).getDate();

  const eventsByDay = useMemo(() => {
    const map = {};
    trainings.forEach((t) => {
      const d = new Date(t.date);
      if (isNaN(d.getTime())) return;
      const key = `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`;
      (map[key] = map[key] || []).push({ type: "training", data: t, date: d });
    });
    matches.forEach((m) => {
      const d = new Date(m.date);
      if (isNaN(d.getTime())) return;
      const key = `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`;
      (map[key] = map[key] || []).push({ type: "match", data: m, date: d });
    });
    return map;
  }, [trainings, matches]);

  const cells = [];
  for (let i = 0; i < startWeekday; i++) cells.push(null);
  for (let day = 1; day <= daysInMonth; day++) cells.push(day);

  const today = new Date();
  const selectedEvents = selectedDay ? eventsByDay[`${selectedDay.getFullYear()}-${selectedDay.getMonth()}-${selectedDay.getDate()}`] || [] : [];

  return (
    <div>
      <SectionTitle
        eyebrow="Agenda"
        title="Calendario"
        icon={Calendar}
        action={
          <div className="flex items-center gap-1.5">
            <button onClick={() => setCursor(new Date(year, month - 1, 1))} className="rounded-lg p-2 hover:bg-white/10 text-slate-400">
              <ChevronLeft className="w-4 h-4" />
            </button>
            <p className="text-sm font-bold text-slate-100 w-36 text-center">{ITALIAN_MONTHS[month]} {year}</p>
            <button onClick={() => setCursor(new Date(year, month + 1, 1))} className="rounded-lg p-2 hover:bg-white/10 text-slate-400">
              <ChevronRight className="w-4 h-4" />
            </button>
            <Button variant="secondary" onClick={() => setCursor(new Date())}>Oggi</Button>
          </div>
        }
      />

      <div className="grid grid-cols-7 gap-1 mb-1">
        {ITALIAN_WEEKDAYS.map((w) => (
          <div key={w} className="text-center text-[10px] font-semibold text-slate-500 uppercase py-0.5">{w}</div>
        ))}
      </div>

      <div className="grid grid-cols-7 gap-1">
        {cells.map((day, i) => {
          if (day === null) return <div key={`empty-${i}`} />;
          const d = new Date(year, month, day);
          const key = `${year}-${month}-${day}`;
          const events = eventsByDay[key] || [];
          const isToday = sameDay(d, today);
          return (
            <button
              key={key}
              onClick={() => setSelectedDay(d)}
              className={`h-12 sm:h-14 rounded-lg border p-1 flex flex-col items-center sm:items-start gap-0.5 text-left transition-colors ${
                isToday ? "border-emerald-500/50 bg-emerald-500/10" : "border-white/10 hover:bg-white/5"
              }`}
            >
              <span className={`text-[11px] font-semibold ${isToday ? "text-emerald-400" : "text-slate-300"}`}>{day}</span>
              <div className="flex flex-wrap gap-0.5 justify-center sm:justify-start">
                {events.some((e) => e.type === "training") && <span className="w-1.5 h-1.5 rounded-full bg-emerald-400" />}
                {events.some((e) => e.type === "match") && <span className="w-1.5 h-1.5 rounded-full bg-sky-400" />}
              </div>
            </button>
          );
        })}
      </div>

      <div className="flex items-center gap-4 mt-4 text-[11px] text-slate-500">
        <span className="flex items-center gap-1.5"><span className="w-2 h-2 rounded-full bg-emerald-400" /> Allenamento</span>
        <span className="flex items-center gap-1.5"><span className="w-2 h-2 rounded-full bg-sky-400" /> Partita</span>
      </div>

      <Modal open={!!selectedDay} onClose={() => setSelectedDay(null)} title={selectedDay ? formatDate(selectedDay.toISOString()) : ""}>
        {selectedEvents.length === 0 ? (
          <EmptyState icon={Calendar} text="Nessun evento in questo giorno." />
        ) : (
          <div className="space-y-2">
            {selectedEvents.map((e, i) => (
              <Card key={i} className="p-3">
                {e.type === "training" ? (
                  <div className="flex items-center justify-between gap-2">
                    <div>
                      <p className="text-sm font-semibold text-emerald-400">Allenamento</p>
                      <p className="text-xs text-slate-400">{e.data.time || "--"} · {e.data.focus || "Nessun focus indicato"}</p>
                    </div>
                    <Button variant="secondary" className="px-2.5 py-1 text-xs" onClick={() => onGoTo("trainings")}>Vai</Button>
                  </div>
                ) : (
                  <div className="flex items-center justify-between gap-2">
                    <div>
                      <p className="text-sm font-semibold text-sky-400">vs {e.data.opponent}</p>
                      <p className="text-xs text-slate-400">{e.data.time || "--"} · {e.data.venue || ""}</p>
                    </div>
                    <Button variant="secondary" className="px-2.5 py-1 text-xs" onClick={() => onGoTo("matches")}>Vai</Button>
                  </div>
                )}
              </Card>
            ))}
          </div>
        )}
      </Modal>
    </div>
  );
}

/* ============================================================
   SEZIONE GIOCATORI
   ============================================================ */

function PlayersSection({ season, updateSeason, showToast, view, setView, jumpToPlayerId, onJumpHandled }) {
  const config = useConfig();
  const players = season.players || [];
  const trainings = season.trainings || [];
  const matches = season.matches || [];
  const [search, setSearch] = useState("");
  const [roleFilter, setRoleFilter] = useState("Tutti");
  const [showAdd, setShowAdd] = useState(false);
  const [selectedPlayer, setSelectedPlayer] = useState(null);

  useEffect(() => {
    if (!jumpToPlayerId) return;
    const p = players.find((pl) => pl.id === jumpToPlayerId);
    if (p) setSelectedPlayer(p);
    onJumpHandled?.();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [jumpToPlayerId]);

  const filtered = players.filter((p) => {
    const matchesSearch = p.name.toLowerCase().includes(search.toLowerCase());
    const matchesRole = roleFilter === "Tutti" || p.role === roleFilter;
    return matchesSearch && matchesRole;
  });

  function addPlayer(playerData) {
    updateSeason((s) => ({ players: [...(s.players || []), { ...playerData, id: uid("player") }] }));
    setShowAdd(false);
    showToast("Giocatore aggiunto alla rosa");
  }

  function updatePlayer(id, patch) {
    updateSeason((s) => ({
      players: (s.players || []).map((p) => (p.id === id ? { ...p, ...patch } : p)),
    }));
  }

  function deletePlayer(id) {
    const removed = players.find((p) => p.id === id);
    updateSeason((s) => ({ players: (s.players || []).filter((p) => p.id !== id) }));
    setSelectedPlayer(null);
    showToast(
      `${removed?.name || "Giocatore"} rimosso dalla rosa`,
      "success",
      removed
        ? { label: "Annulla", onClick: () => updateSeason((s) => ({ players: [...(s.players || []), removed] })) }
        : null
    );
  }

  const [showUnavailable, setShowUnavailable] = useState(false);
  const [showCompare, setShowCompare] = useState(false);
  const unavailablePlayers = players.filter((p) => ["Infortunato", "Squalificato", "In dubbio"].includes(p.medicalStatus));

  return (
    <div>
      <SectionTitle
        eyebrow="Anagrafica"
        title={`Giocatori (${players.length}/30)`}
        icon={Users}
        action={
          <div className="flex gap-2 flex-wrap">
            <div className="flex rounded-xl border border-white/10 overflow-hidden">
              <button
                onClick={() => setView("grid")}
                className={`px-3 py-2 text-xs font-medium ${view === "grid" ? "bg-emerald-500 text-slate-950" : "text-slate-400"}`}
              >
                Griglia
              </button>
              <button
                onClick={() => setView("board")}
                className={`px-3 py-2 text-xs font-medium ${view === "board" ? "bg-emerald-500 text-slate-950" : "text-slate-400"}`}
              >
                Tabellone
              </button>
            </div>
            <button
              onClick={() => setShowUnavailable(true)}
              className="rounded-xl text-white px-4 py-2.5 text-sm font-semibold flex items-center gap-2 transition-opacity hover:opacity-90"
              style={{ backgroundColor: "#7a2333" }}
            >
              <AlertCircle className="w-4 h-4" /> Indisponibili ({unavailablePlayers.length})
            </button>
            <Button variant="secondary" onClick={() => setShowCompare(true)} disabled={players.length < 2}>
              <ArrowLeftRight className="w-4 h-4" /> Confronta
            </Button>
            <Button onClick={() => setShowAdd(true)}>
              <Plus className="w-4 h-4" /> Aggiungi Giocatore
            </Button>
            <SectionResetButton
              label="Azzera Giocatori"
              confirmText="Eliminare tutti i giocatori della rosa?"
              onConfirm={() => {
                updateSeason(() => ({ players: [] }));
                showToast("Rosa giocatori azzerata");
              }}
            />
          </div>
        }
      />

      <div className="flex flex-col sm:flex-row gap-3 mb-5">
        <div className="relative flex-1">
          <Search className="w-4 h-4 text-slate-500 absolute left-3 top-1/2 -translate-y-1/2" />
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Cerca giocatore..."
            className={inputClass + " pl-9"}
          />
        </div>
        <div className="flex items-center gap-2 overflow-x-auto pb-1">
          {["Tutti", ...config.roles].map((r) => (
            <button
              key={r}
              onClick={() => setRoleFilter(r)}
              className={`shrink-0 rounded-full px-3.5 py-2 text-xs font-medium border transition-colors ${
                roleFilter === r
                  ? "bg-emerald-500 text-slate-950 border-emerald-500"
                  : "border-white/10 text-slate-400 hover:text-slate-200"
              }`}
            >
              {r}
            </button>
          ))}
        </div>
      </div>

      {view === "board" ? (
        <PlayersBoard players={filtered} trainings={trainings} matches={matches} onSelect={setSelectedPlayer} />
      ) : filtered.length === 0 ? (
        <EmptyState icon={Users} text="Nessun giocatore trovato. Aggiungi il primo giocatore alla rosa." />
      ) : (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: "16px" }}>
          {filtered.map((p) => (
            <Card key={p.id} className="p-3 sm:p-4 hover:border-emerald-500/40 transition-colors h-full flex flex-col items-center text-center">
              <button onClick={() => setSelectedPlayer(p)} className="flex flex-col items-center">
                <div className="relative mb-2 sm:mb-3">
                  <img src={playerAvatar(p)} alt={p.name} className="w-14 h-14 sm:w-16 sm:h-16 rounded-full bg-slate-800 border-2 border-white/10 object-cover" />
                  <span className="absolute -bottom-1 -right-1 w-6 h-6 rounded-full bg-emerald-500 text-slate-950 text-[11px] font-extrabold flex items-center justify-center border-2 border-slate-900">
                    {p.number ?? "-"}
                  </span>
                </div>
                <p className="text-xs sm:text-sm font-bold text-slate-100 truncate w-full">{p.name}</p>
              </button>

              <div className="mt-2 flex items-center gap-1 flex-wrap justify-center">
                <InlinePlayerSelect
                  value={p.role}
                  options={config.roles}
                  onChange={(v) => updatePlayer(p.id, { role: v })}
                  className="border-emerald-500/30 text-emerald-400 bg-emerald-500/10"
                />
                {p.role !== "Portiere" && (
                  <InlinePlayerSelect
                    value={p.position || config.positions[0] || "Centro"}
                    options={config.positions}
                    onChange={(v) => updatePlayer(p.id, { position: v })}
                    className="border-emerald-500/30 text-emerald-400 bg-emerald-500/10"
                  />
                )}
              </div>

              <div className="mt-1.5 flex items-center gap-1 flex-wrap justify-center">
                <InlinePlayerSelect
                  value={p.role2 || ""}
                  options={[{ value: "", label: "Alt: —" }, ...config.roles.filter((r) => r !== p.role).map((r) => ({ value: r, label: `Alt: ${r}` }))]}
                  onChange={(v) => updatePlayer(p.id, { role2: v || null })}
                  className="border-white/20 text-slate-200 bg-white/5"
                />
                <InlinePlayerSelect
                  value={p.medicalStatus}
                  options={config.medicalStatuses}
                  onChange={(v) => updatePlayer(p.id, { medicalStatus: v })}
                  className={MEDICAL_COLORS[p.medicalStatus] || NEUTRAL_BADGE}
                />
              </div>
            </Card>
          ))}
        </div>
      )}

      <Modal open={showAdd} onClose={() => setShowAdd(false)} title="Aggiungi Giocatore" wide>
        <PlayerForm onSubmit={addPlayer} onCancel={() => setShowAdd(false)} />
      </Modal>

      <Modal open={!!selectedPlayer} onClose={() => setSelectedPlayer(null)} title={selectedPlayer?.name || ""} wide>
        {selectedPlayer && (
          <PlayerDetail
            player={players.find((p) => p.id === selectedPlayer.id) || selectedPlayer}
            trainings={trainings}
            matches={matches}
            onUpdate={(patch) => {
              updatePlayer(selectedPlayer.id, patch);
              setSelectedPlayer((prev) => ({ ...prev, ...patch }));
            }}
            onDelete={() => deletePlayer(selectedPlayer.id)}
            onClose={() => setSelectedPlayer(null)}
          />
        )}
      </Modal>

      <Modal open={showUnavailable} onClose={() => setShowUnavailable(false)} title="Giocatori Indisponibili" wide>
        {unavailablePlayers.length === 0 ? (
          <EmptyState icon={CheckCircle2} text="Nessun giocatore indisponibile al momento: rosa al completo." />
        ) : (
          <div className="overflow-x-auto rounded-2xl border border-white/10">
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-slate-900/80 text-left text-[11px] uppercase tracking-wide text-slate-500">
                  <th className="px-3 py-2.5">#</th>
                  <th className="px-3 py-2.5">Giocatore</th>
                  <th className="px-3 py-2.5">Ruolo</th>
                  <th className="px-3 py-2.5">Stato</th>
                </tr>
              </thead>
              <tbody>
                {[...unavailablePlayers]
                  .sort((a, b) => {
                    const order = { "In dubbio": 0, Infortunato: 1, Squalificato: 2 };
                    return order[a.medicalStatus] - order[b.medicalStatus] || a.name.localeCompare(b.name);
                  })
                  .map((p) => {
                    const inDoubt = p.medicalStatus === "In dubbio";
                    return (
                      <tr
                        key={p.id}
                        onClick={() => {
                          setShowUnavailable(false);
                          setSelectedPlayer(p);
                        }}
                        className={`border-t border-white/5 cursor-pointer hover:bg-white/5 ${inDoubt ? "bg-rose-500/10" : ""}`}
                      >
                        <td className="px-3 py-2.5 font-bold text-emerald-400">{p.number ?? "-"}</td>
                        <td className="px-3 py-2.5">
                          <div className="flex items-center gap-2">
                            <img src={playerAvatar(p)} alt={p.name} className="w-7 h-7 rounded-full bg-slate-800 object-cover shrink-0" />
                            <span className={`font-medium truncate ${inDoubt ? "text-rose-300" : "text-slate-200"}`}>{p.name}</span>
                          </div>
                        </td>
                        <td className="px-3 py-2.5">
                          <Badge className={ROLE_COLORS[p.role]}>{p.role}</Badge>
                        </td>
                        <td className="px-3 py-2.5">
                          <Badge className={inDoubt ? "bg-rose-500/20 text-rose-300 border-rose-500/40" : MEDICAL_COLORS[p.medicalStatus]}>
                            <HeartPulse className="w-3 h-3" /> {p.medicalStatus}
                          </Badge>
                        </td>
                      </tr>
                    );
                  })}
              </tbody>
            </table>
          </div>
        )}
      </Modal>

      <Modal open={showCompare} onClose={() => setShowCompare(false)} title="Confronta Giocatori" wide>
        <PlayerCompareModal players={players} onClose={() => setShowCompare(false)} />
      </Modal>
    </div>
  );
}

function PlayersBoard({ players, trainings, matches, onSelect }) {
  if (players.length === 0) {
    return <EmptyState icon={Users} text="Nessun giocatore trovato." />;
  }
  const sorted = [...players].sort((a, b) => (a.number ?? 999) - (b.number ?? 999));
  return (
    <div className="overflow-x-auto rounded-2xl border border-white/10">
      <table className="w-full text-sm">
        <thead>
          <tr className="bg-slate-900/80 text-left text-[11px] uppercase tracking-wide text-slate-500">
            <th className="px-3 py-2.5">#</th>
            <th className="px-3 py-2.5">Giocatore</th>
            <th className="px-3 py-2.5 text-center bg-emerald-500/10 text-emerald-400">Presenze All.</th>
            <th className="px-3 py-2.5 text-center bg-rose-500/10 text-rose-400">Assenze All.</th>
            <th className="px-3 py-2.5 text-center">Convocazioni</th>
            <th className="px-3 py-2.5 text-center">Reti</th>
            <th className="px-3 py-2.5 text-center">Assist</th>
            <th className="px-3 py-2.5 text-center">Amm.</th>
            <th className="px-3 py-2.5 text-center">Esp.</th>
          </tr>
        </thead>
        <tbody>
          {sorted.map((p) => {
            const s = computePlayerStats(p.id, trainings, matches);
            return (
              <tr key={p.id} onClick={() => onSelect(p)} className="border-t border-white/5 hover:bg-white/5 cursor-pointer">
                <td className="px-3 py-2.5 font-bold text-emerald-400">{p.number ?? "-"}</td>
                <td className="px-3 py-2.5">
                  <div className="flex items-center gap-2">
                    <img src={playerAvatar(p)} alt={p.name} className="w-7 h-7 rounded-full bg-slate-800 object-cover shrink-0" />
                    <div className="min-w-0">
                      <p className="text-slate-200 font-medium truncate">{p.name}</p>
                      <p className="text-[10px] text-slate-500">{p.role}</p>
                    </div>
                  </div>
                </td>
                <td className="px-3 py-2.5 text-center bg-emerald-500/10 text-emerald-400 font-bold">{s.presenze}</td>
                <td className="px-3 py-2.5 text-center bg-rose-500/10 text-rose-400 font-bold">{s.assenze}</td>
                <td className="px-3 py-2.5 text-center text-sky-400">{s.convocazioni}</td>
                <td className="px-3 py-2.5 text-center text-slate-200 font-semibold">{s.reti}</td>
                <td className="px-3 py-2.5 text-center text-slate-200 font-semibold">{s.assist}</td>
                <td className="px-3 py-2.5 text-center text-amber-400">{s.ammonizioni}</td>
                <td className="px-3 py-2.5 text-center text-rose-500">{s.espulsioni}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

const COMPARE_COLORS = ["#10b981", "#38bdf8", "#f472b6"];

const COMPARE_STAT_GROUPS = [
  { id: "base", label: "Caratteristiche base", keys: BASE_STAT_KEYS, field: "baseStats" },
  { id: "mentali", label: "Mentali", keys: MENTAL_STAT_KEYS, field: "mentalStats" },
  { id: "tecnico", label: "Tecnico/Tattiche", keys: TECH_TACTIC_STAT_KEYS, field: "techTacticStats" },
  { id: "portiere", label: "Portiere", keys: GK_STAT_KEYS, field: "gkStats" },
];

// Confronto tra 2-3 giocatori con radar sovrapposti: un solo grafico con una
// serie colorata per giocatore, per decisioni tattiche a colpo d'occhio.
function PlayerCompareModal({ players }) {
  const [selectedIds, setSelectedIds] = useState([]);
  const [groupId, setGroupId] = useState("base");

  const selectedPlayers = selectedIds.map((id) => players.find((p) => p.id === id)).filter(Boolean);
  const group = COMPARE_STAT_GROUPS.find((g) => g.id === groupId) || COMPARE_STAT_GROUPS[0];

  function toggleSelect(id) {
    setSelectedIds((prev) => {
      if (prev.includes(id)) return prev.filter((x) => x !== id);
      if (prev.length >= 3) return prev; // massimo 3 giocatori confrontabili, per leggibilità del grafico
      return [...prev, id];
    });
  }

  const radarData = group.keys.map((s) => {
    const row = { stat: s.label };
    selectedPlayers.forEach((p) => {
      const statsObj = p[group.field] || {};
      row[p.id] = clampStat(statsObj[s.key]);
    });
    return row;
  });

  return (
    <div>
      <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide mb-2">Scegli 2 o 3 giocatori</p>
      <div className="flex flex-wrap gap-1.5 mb-4 max-h-40 overflow-y-auto pr-1">
        {players.map((p) => {
          const isSelected = selectedIds.includes(p.id);
          const colorIdx = selectedIds.indexOf(p.id);
          return (
            <button
              key={p.id}
              onClick={() => toggleSelect(p.id)}
              disabled={!isSelected && selectedIds.length >= 3}
              className={`flex items-center gap-1.5 rounded-full pl-1 pr-3 py-1 text-xs font-medium border transition-colors disabled:opacity-40 ${
                isSelected ? "border-transparent text-slate-950" : "border-white/10 text-slate-300 hover:text-slate-100"
              }`}
              style={isSelected ? { backgroundColor: COMPARE_COLORS[colorIdx] } : {}}
            >
              <img src={playerAvatar(p)} alt={p.name} className="w-5 h-5 rounded-full object-cover" />
              {p.name}
            </button>
          );
        })}
      </div>

      <div className="flex flex-wrap gap-1.5 mb-4">
        {COMPARE_STAT_GROUPS.map((g) => (
          <button
            key={g.id}
            onClick={() => setGroupId(g.id)}
            className={`rounded-full px-3 py-1.5 text-[11px] font-medium border transition-colors ${
              groupId === g.id ? "bg-sky-500 text-slate-950 border-sky-500" : "border-white/10 text-slate-400 hover:text-slate-200"
            }`}
          >
            {g.label}
          </button>
        ))}
      </div>

      {selectedPlayers.length < 2 ? (
        <EmptyState icon={ArrowLeftRight} text="Seleziona almeno 2 giocatori per vedere il confronto." />
      ) : (
        <div className="h-80">
          <ResponsiveContainer width="100%" height="100%">
            <RadarChart data={radarData}>
              <PolarGrid stroke="rgba(255,255,255,0.1)" />
              <PolarAngleAxis dataKey="stat" tick={{ fill: "#94a3b8", fontSize: 10 }} />
              <PolarRadiusAxis domain={[0, 10]} tick={false} axisLine={false} />
              {selectedPlayers.map((p, i) => (
                <Radar key={p.id} name={p.name} dataKey={p.id} stroke={COMPARE_COLORS[i]} fill={COMPARE_COLORS[i]} fillOpacity={0.15} />
              ))}
              <Legend />
            </RadarChart>
          </ResponsiveContainer>
        </div>
      )}
    </div>
  );
}

function PlayerForm({ onSubmit, onCancel, initial }) {
  const config = useConfig();
  const [form, setForm] = useState(
    initial || {
      name: "",
      role: config.roles[2] || config.roles[0] || "Centrocampista",
      role2: "",
      position: config.positions[0] || "Centro",
      preferredFoot: "Destro",
      height: "",
      number: "",
      medicalStatus: config.medicalStatuses[0] || "Disponibile",
      birthDate: "",
      notes: "",
      photo: null,
      baseStats: emptyBaseStats(),
      mentalStats: emptyMentalStats(),
      techTacticStats: emptyTechTacticStats(),
      gkStats: emptyGkStats(),
      coachNotes: emptyCoachNotes(),
    }
  );
  const [photoError, setPhotoError] = useState("");
  const fileInputRef = React.useRef(null);
  const isGoalkeeper = form.role === "Portiere";

  async function handlePhotoChange(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    if (!file.type.startsWith("image/")) {
      setPhotoError("Seleziona un file immagine valido");
      return;
    }
    try {
      setPhotoError("");
      const dataUrl = await resizeImageFile(file);
      setForm((f) => ({ ...f, photo: dataUrl }));
    } catch (err) {
      setPhotoError("Impossibile caricare l'immagine");
    }
    e.target.value = "";
  }

  return (
    <div>
      <div className="flex items-center gap-4 mb-5">
        <img
          src={form.photo || avatarUrl(form.name || "player")}
          alt="Anteprima"
          className="w-20 h-20 rounded-2xl bg-slate-800 border-2 border-white/10 object-cover"
        />
        <div>
          <div className="flex gap-2">
            <Button type="button" variant="secondary" onClick={() => fileInputRef.current?.click()}>
              Carica foto
            </Button>
            {form.photo && (
              <Button type="button" variant="ghost" onClick={() => setForm((f) => ({ ...f, photo: null }))}>
                Rimuovi
              </Button>
            )}
          </div>
          <input ref={fileInputRef} type="file" accept="image/*" className="hidden" onChange={handlePhotoChange} />
          <p className="text-[11px] text-slate-500 mt-1.5">Senza foto viene usato un avatar generato automaticamente.</p>
          {photoError && <p className="text-[11px] text-rose-400 mt-1">{photoError}</p>}
        </div>
      </div>

      <div className="grid sm:grid-cols-2 gap-x-4">
        <Field label="Nome e Cognome">
          <input className={inputClass} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="Es. Marco Rossi" />
        </Field>
        <Field label="Numero di maglia (facoltativo)">
          <input type="number" className={inputClass} value={form.number} onChange={(e) => setForm({ ...form, number: e.target.value })} placeholder="Lascia vuoto se non assegnato" />
        </Field>
        <Field label="Ruolo principale">
          <select className={inputClass} value={form.role} onChange={(e) => setForm({ ...form, role: e.target.value })}>
            {config.roles.map((r) => (
              <option key={r} value={r}>{r}</option>
            ))}
          </select>
        </Field>
        <Field label="Ruolo alternativo">
          <select className={inputClass} value={form.role2 || ""} onChange={(e) => setForm({ ...form, role2: e.target.value })}>
            <option value="">Nessuno</option>
            {config.roles.filter((r) => r !== form.role).map((r) => (
              <option key={r} value={r}>{r}</option>
            ))}
          </select>
        </Field>
        {!isGoalkeeper && (
          <Field label="Posizione in campo">
            <select className={inputClass} value={form.position || config.positions[0] || "Centro"} onChange={(e) => setForm({ ...form, position: e.target.value })}>
              {config.positions.map((p) => (
                <option key={p} value={p}>{p}</option>
              ))}
            </select>
          </Field>
        )}
        <Field label="Piede preferito">
          <select className={inputClass} value={form.preferredFoot || "Destro"} onChange={(e) => setForm({ ...form, preferredFoot: e.target.value })}>
            {FEET.map((f) => (
              <option key={f} value={f}>{f}</option>
            ))}
          </select>
        </Field>
        <Field label="Altezza (cm)">
          <input type="number" className={inputClass} value={form.height} onChange={(e) => setForm({ ...form, height: e.target.value })} placeholder="Es. 165" />
        </Field>
        <Field label="Stato medico">
          <select className={inputClass} value={form.medicalStatus} onChange={(e) => setForm({ ...form, medicalStatus: e.target.value })}>
            {config.medicalStatuses.map((m) => (
              <option key={m} value={m}>{m}</option>
            ))}
          </select>
        </Field>
        <Field label="Data di nascita">
          <input type="date" className={inputClass} value={form.birthDate} onChange={(e) => setForm({ ...form, birthDate: e.target.value })} />
        </Field>
        <Field label="Note">
          <input className={inputClass} value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} placeholder="Facoltativo" />
        </Field>
      </div>

      <div className="flex justify-end gap-2 mt-4">
        <Button variant="secondary" onClick={onCancel}>Annulla</Button>
        <Button
          onClick={() => {
            if (!form.name.trim()) return;
            onSubmit({
              ...form,
              number: form.number ? Number(form.number) : null,
              height: form.height ? Number(form.height) : null,
              position: isGoalkeeper ? null : form.position || "Centro",
              avatarSeed: form.name,
              baseStats: form.baseStats || emptyBaseStats(),
              mentalStats: form.mentalStats || emptyMentalStats(),
              techTacticStats: form.techTacticStats || emptyTechTacticStats(),
              gkStats: form.gkStats || emptyGkStats(),
              coachNotes: form.coachNotes || emptyCoachNotes(),
            });
          }}
        >
          <Save className="w-4 h-4" /> Salva Giocatore
        </Button>
      </div>
    </div>
  );
}

function PlayerDetail({ player, onUpdate, onDelete, onClose, trainings, matches }) {
  const [tab, setTab] = useState("overview");
  const [editing, setEditing] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);

  const isGoalkeeper = player.role === "Portiere";
  const baseStats = player.baseStats || emptyBaseStats();
  const mentalStats = player.mentalStats || emptyMentalStats();
  const techTacticStats = player.techTacticStats || emptyTechTacticStats();
  const gkStats = player.gkStats || emptyGkStats();
  const coachNotes = player.coachNotes || emptyCoachNotes();
  const stats = computePlayerStats(player.id, trainings, matches);

  const radarData = BASE_STAT_KEYS.map((s) => ({ stat: s.label, value: clampStat(baseStats[s.key]) }));
  const mentalRadarData = MENTAL_STAT_KEYS.map((s) => ({ stat: s.label, value: clampStat(mentalStats[s.key]) }));
  const techTacticRadarData = TECH_TACTIC_STAT_KEYS.map((s) => ({ stat: s.label, value: clampStat(techTacticStats[s.key]) }));
  const gkRadarData = GK_STAT_KEYS.map((s) => ({ stat: s.label, value: clampStat(gkStats[s.key]) }));

  if (editing) {
    return (
      <PlayerForm
        initial={{ ...player }}
        onCancel={() => setEditing(false)}
        onSubmit={(patch) => {
          onUpdate(patch);
          setEditing(false);
        }}
      />
    );
  }

  function downloadPlayerSheet() {
    let body = `<h1>${player.name} #${player.number ?? "-"}</h1>`;
    body += `<p>${player.role}${player.role2 ? ` · Alt: ${player.role2}` : ""} · ${player.medicalStatus}</p>`;
    const infoBits = [];
    if (!isGoalkeeper && player.position) infoBits.push(`Posizione: ${player.position}`);
    if (player.preferredFoot) infoBits.push(`Piede: ${player.preferredFoot}`);
    if (player.height) infoBits.push(`Altezza: ${player.height} cm`);
    if (player.birthDate) infoBits.push(`Nato il: ${formatDate(player.birthDate)}`);
    if (player.notes) infoBits.push(`Note: ${player.notes}`);
    if (infoBits.length) body += `<p>${infoBits.join(" · ")}</p>`;

    body += `<h2>Statistiche stagione</h2><p>Presenze: ${stats.presenze} · Assenze: ${stats.assenze} · Convocazioni: ${stats.convocazioni} · Reti: ${stats.reti} · Assist: ${stats.assist} · Ammonizioni: ${stats.ammonizioni} · Espulsioni: ${stats.espulsioni}</p>`;

    body += `<h2>Caratteristiche base</h2><table><tr>${BASE_STAT_KEYS.map((s) => `<th>${s.label}</th>`).join("")}</tr><tr>${BASE_STAT_KEYS.map((s) => `<td>${baseStats[s.key] ?? 5}</td>`).join("")}</tr></table>`;

    body += `<h2>Statistiche Mentali</h2><table><tr>${MENTAL_STAT_KEYS.map((s) => `<th>${s.label}</th>`).join("")}</tr><tr>${MENTAL_STAT_KEYS.map((s) => `<td>${mentalStats[s.key] ?? 5}</td>`).join("")}</tr></table>`;

    if (isGoalkeeper) {
      body += `<h2>Statistiche Portiere</h2><table><tr>${GK_STAT_KEYS.map((s) => `<th>${s.label}</th>`).join("")}</tr><tr>${GK_STAT_KEYS.map((s) => `<td>${gkStats[s.key] ?? 5}</td>`).join("")}</tr></table>`;
    } else {
      body += `<h2>Tecnico/Tattiche</h2><table><tr>${TECH_TACTIC_STAT_KEYS.map((s) => `<th>${s.label}</th>`).join("")}</tr><tr>${TECH_TACTIC_STAT_KEYS.map((s) => `<td>${techTacticStats[s.key] ?? 5}</td>`).join("")}</tr></table>`;
    }

    body += `<h2>Valutazione Mister</h2>`;
    BASE_STAT_KEYS.forEach((s) => {
      if (coachNotes[s.key]) body += `<p><strong>${s.label}:</strong> ${coachNotes[s.key]}</p>`;
    });

    downloadPrintableHTML(`scheda-${player.name}.html`, `Scheda Giocatore — ${player.name}`, body);
  }

  const TABS = [
    { id: "overview", label: "Caratteristiche base" },
    { id: "mental", label: "Statistiche Mentali" },
    isGoalkeeper ? { id: "gk", label: "Statistiche Portiere" } : { id: "techtactic", label: "Tecnico/Tattiche" },
    { id: "coach", label: "Valutazione Mister" },
  ];

  return (
    <div>
      <div className="flex items-center gap-4 mb-3">
        <img src={playerAvatar(player)} alt={player.name} className="w-20 h-20 rounded-2xl bg-slate-800 border-2 border-white/10 object-cover" />
        <div className="flex-1">
          <div className="flex items-center gap-2 flex-wrap">
            <h4 className="text-lg font-bold text-slate-100">{player.name}</h4>
            <span className="text-emerald-400 font-extrabold">#{player.number ?? "-"}</span>
          </div>
          <div className="flex items-center gap-2 mt-1.5 flex-wrap">
            <Badge className={ROLE_COLORS[player.role]}>{player.role}</Badge>
            {player.role2 && <Badge className="bg-white/5 text-slate-300 border-white/10">Alt: {player.role2}</Badge>}
            <Badge className={MEDICAL_COLORS[player.medicalStatus]}>
              <HeartPulse className="w-3 h-3" /> {player.medicalStatus}
            </Badge>
          </div>
        </div>
        <div className="flex gap-1.5">
          <button onClick={downloadPlayerSheet} className="rounded-lg p-2 hover:bg-white/10 text-slate-400" title="Scarica scheda giocatore">
            <Download className="w-4 h-4" />
          </button>
          <button onClick={() => setEditing(true)} className="rounded-lg p-2 hover:bg-white/10 text-slate-400" title="Modifica">
            <Edit2 className="w-4 h-4" />
          </button>
          <button onClick={() => setConfirmDelete(true)} className="rounded-lg p-2 hover:bg-rose-500/10 text-rose-400" title="Elimina">
            <Trash2 className="w-4 h-4" />
          </button>
          {onClose && (
            <button onClick={onClose} className="rounded-lg p-2 hover:bg-white/10 text-slate-400" title="Torna al menu">
              <ChevronLeft className="w-4 h-4" />
            </button>
          )}
        </div>
      </div>

      <div className="flex flex-wrap gap-x-4 gap-y-1 text-xs text-slate-400 mb-5">
        {!isGoalkeeper && player.position && <span>Posizione: <span className="text-slate-300">{player.position}</span></span>}
        {player.preferredFoot && <span>Piede: <span className="text-slate-300">{player.preferredFoot}</span></span>}
        {player.height && <span>Altezza: <span className="text-slate-300">{player.height} cm</span></span>}
        {player.birthDate && <span>Nato il: <span className="text-slate-300">{formatDate(player.birthDate)}</span></span>}
        {player.notes && <span>Note: <span className="text-slate-300">{player.notes}</span></span>}
      </div>

      <div className="flex gap-1 mb-5 overflow-x-auto pb-0.5">
        {[
          { label: "Presenze", value: stats.presenze, color: "text-emerald-400" },
          { label: "Assenze", value: stats.assenze, color: "text-rose-400" },
          { label: "Convoc.", value: stats.convocazioni, color: "text-sky-400" },
          { label: "Reti", value: stats.reti, color: "text-slate-100" },
          { label: "Assist", value: stats.assist, color: "text-slate-100" },
          { label: "Amm.", value: stats.ammonizioni, color: "text-amber-400" },
          { label: "Esp.", value: stats.espulsioni, color: "text-rose-500" },
        ].map((k) => (
          <div key={k.label} className="flex-1 rounded-lg border border-white/10 p-1 text-center" style={{ minWidth: 52 }}>
            <p className={`text-xs font-extrabold leading-tight ${k.color}`}>{k.value}</p>
            <p className="text-[8px] text-slate-500 leading-tight whitespace-nowrap">{k.label}</p>
          </div>
        ))}
      </div>

      {confirmDelete && (
        <div className="rounded-xl border border-rose-500/30 bg-rose-500/10 p-3 mb-4 flex items-center justify-between gap-2">
          <p className="text-sm text-rose-300">Rimuovere definitivamente questo giocatore dalla rosa?</p>
          <div className="flex gap-2 shrink-0">
            <Button variant="secondary" onClick={() => setConfirmDelete(false)}>Annulla</Button>
            <Button variant="danger" onClick={onDelete}>Elimina</Button>
          </div>
        </div>
      )}

      <div className="flex gap-1 mb-3 border-b border-white/10 overflow-x-auto">
        {TABS.map((t) => (
          <button
            key={t.id}
            onClick={() => setTab(t.id)}
            className={`px-3 py-2.5 text-sm font-medium border-b-2 -mb-px whitespace-nowrap ${
              tab === t.id ? "border-emerald-400 text-emerald-400" : "border-transparent text-slate-500"
            }`}
          >
            {t.label}
          </button>
        ))}
      </div>

      <div
        className="overflow-y-auto pr-2 -mr-2 border border-white/5 rounded-xl p-3 bg-slate-950/30"
        style={{ maxHeight: "min(58vh, 560px)" }}
      >
        {tab === "overview" ? (
          <StatsPanel
            radarData={radarData}
            radarColor="#10b981"
            statKeys={BASE_STAT_KEYS}
            values={baseStats}
            onChange={(key, v) => onUpdate({ baseStats: { ...baseStats, [key]: v } })}
          />
        ) : tab === "mental" ? (
          <StatsPanel
            radarData={mentalRadarData}
            radarColor="#a855f7"
            statKeys={MENTAL_STAT_KEYS}
            values={mentalStats}
            onChange={(key, v) => onUpdate({ mentalStats: { ...mentalStats, [key]: v } })}
          />
        ) : tab === "techtactic" ? (
          <StatsPanel
            radarData={techTacticRadarData}
            radarColor="#0ea5e9"
            statKeys={TECH_TACTIC_STAT_KEYS}
            values={techTacticStats}
            onChange={(key, v) => onUpdate({ techTacticStats: { ...techTacticStats, [key]: v } })}
          />
        ) : tab === "gk" ? (
          <StatsPanel
            radarData={gkRadarData}
            radarColor="#f59e0b"
            statKeys={GK_STAT_KEYS}
            values={gkStats}
            onChange={(key, v) => onUpdate({ gkStats: { ...gkStats, [key]: v } })}
          />
        ) : (
          <div>
            <p className="text-xs text-slate-500 mb-4">
              Valutazione descrittiva del mister per ciascuna caratteristica principale del giocatore.
            </p>
            {BASE_STAT_KEYS.map((s) => (
              <Field key={s.key} label={s.label}>
                <textarea
                  rows={3}
                  className={inputClass}
                  value={coachNotes[s.key] || ""}
                  onChange={(e) => onUpdate({ coachNotes: { ...coachNotes, [s.key]: e.target.value } })}
                  placeholder={`Note del mister su ${s.label.toLowerCase()}...`}
                />
              </Field>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

// Layout compatto: radar a sinistra (larghezza fissa), statistiche a destra su 2-3 colonne
function StatsPanel({ radarData, radarColor, statKeys, values, onChange }) {
  return (
    <div style={{ display: "flex", flexWrap: "wrap", gap: "20px", alignItems: "flex-start" }}>
      <div style={{ flex: "0 0 220px", maxWidth: "260px", margin: "0 auto" }}>
        <ResponsiveContainer width="100%" height={220}>
          <RadarChart data={radarData}>
            <PolarGrid stroke="#ffffff15" />
            <PolarAngleAxis dataKey="stat" tick={{ fill: "#94a3b8", fontSize: 9 }} />
            <PolarRadiusAxis angle={30} domain={[0, STAT_MAX]} tick={{ fill: "#475569", fontSize: 8 }} />
            <Radar dataKey="value" stroke={radarColor} fill={radarColor} fillOpacity={0.3} />
          </RadarChart>
        </ResponsiveContainer>
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-x-4" style={{ flex: "1 1 280px", minWidth: "0" }}>
        {statKeys.map((s) => (
          <StatBar key={s.key} label={s.label} value={clampStat(values[s.key])} editable onChange={(v) => onChange(s.key, v)} />
        ))}
      </div>
    </div>
  );
}

function StatBar({ label, value, editable, onChange }) {
  if (editable) {
    return <Slider label={label} value={value} onChange={onChange} />;
  }
  return (
    <div className="mb-2.5">
      <div className="flex justify-between text-xs mb-1">
        <span className="text-slate-400">{label}</span>
        <span className="text-slate-300 font-semibold">{value}</span>
      </div>
      <div className="h-1.5 rounded-full bg-slate-800 overflow-hidden">
        <div className="h-full bg-emerald-500 rounded-full" style={{ width: `${value}%` }} />
      </div>
    </div>
  );
}

/* ============================================================
   SEZIONE ALLENAMENTI
   ============================================================ */

function TrainingsSection({ season, updateSeason, library, updateLibrary, showToast }) {
  const trainings = [...(season.trainings || [])].sort((a, b) => new Date(b.date) - new Date(a.date));
  const players = season.players || [];
  const focusTecnici = season.focusTecnici || [];
  const [subTab, setSubTab] = useState("sessioni");
  const [showAdd, setShowAdd] = useState(false);
  const [selected, setSelected] = useState(null);

  function addTraining(data) {
    // Tutti i giocatori vengono impostati di default come "Presente"
    const attendance = {};
    players.forEach((p) => (attendance[p.id] = "Presente"));
    updateSeason((s) => ({ trainings: [...(s.trainings || []), { ...data, id: uid("training"), attendance }] }));
    setShowAdd(false);
    showToast("Allenamento creato — presenze impostate su Presente per tutti");
  }

  function updateTraining(id, patch) {
    updateSeason((s) => ({
      trainings: (s.trainings || []).map((t) => (t.id === id ? { ...t, ...patch } : t)),
    }));
  }

  function deleteTraining(id) {
    updateSeason((s) => ({ trainings: (s.trainings || []).filter((t) => t.id !== id) }));
    setSelected(null);
  }

  function saveFocusTecnico(ft) {
    updateSeason((s) => {
      const list = s.focusTecnici || [];
      const exists = list.some((f) => f.id === ft.id);
      const focusTecnici = exists ? list.map((f) => (f.id === ft.id ? ft : f)) : [...list, ft];
      return { focusTecnici };
    });
    showToast("Focus Tecnico salvato");
  }

  function deleteFocusTecnico(id) {
    const removed = focusTecnici.find((f) => f.id === id);
    updateSeason((s) => ({ focusTecnici: (s.focusTecnici || []).filter((f) => f.id !== id) }));
    showToast(
      `Focus Tecnico "${removed?.title || ""}" eliminato`,
      "success",
      removed
        ? { label: "Annulla", onClick: () => updateSeason((s) => ({ focusTecnici: [...(s.focusTecnici || []), removed] })) }
        : null
    );
  }

  function saveExercise(ex) {
    updateLibrary((lib) => {
      const list = lib.exercises || [];
      const exists = list.some((e) => e.id === ex.id);
      return {
        exercises: exists ? list.map((e) => (e.id === ex.id ? ex : e)) : [...list, ex],
      };
    });
    showToast("Esercizio salvato");
  }

  function deleteExercise(id) {
    const removed = (library.exercises || []).find((e) => e.id === id);
    updateLibrary((lib) => ({ exercises: (lib.exercises || []).filter((e) => e.id !== id) }));
    showToast(
      `Esercizio "${removed?.title || ""}" eliminato`,
      "success",
      removed
        ? { label: "Annulla", onClick: () => updateLibrary((lib) => ({ exercises: [...(lib.exercises || []), removed] })) }
        : null
    );
  }

  function updateFocusExercise(focusId, exerciseId, updatedExercise) {
    updateSeason((s) => ({
      focusTecnici: (s.focusTecnici || []).map((ft) =>
        ft.id === focusId
          ? { ...ft, exercises: (ft.exercises || []).map((ex) => (ex.id === exerciseId ? updatedExercise : ex)) }
          : ft
      ),
    }));
    showToast("Esercizio aggiornato nel Focus Tecnico");
  }

  function deleteFocusExercise(focusId, exerciseId) {
    updateSeason((s) => ({
      focusTecnici: (s.focusTecnici || []).map((ft) =>
        ft.id === focusId ? { ...ft, exercises: (ft.exercises || []).filter((ex) => ex.id !== exerciseId) } : ft
      ),
    }));
    showToast("Esercizio eliminato dal Focus Tecnico");
  }

  // Usati dal PlayBook: aggiungere un esercizio a un Focus già esistente
  // (se ha ancora spazio libero, max 8) oppure crearne uno nuovo al volo.
  function addExerciseToFocus(focusId, ex) {
    const target = focusTecnici.find((f) => f.id === focusId);
    if (!target || (target.exercises || []).length >= 8) return;
    saveFocusTecnico({ ...target, exercises: [...(target.exercises || []), { ...ex, id: uid("ex"), sourceExerciseId: ex.id }] });
    showToast(`Aggiunto a "${target.title}"`);
  }

  function createFocusWithExercise(ex) {
    const newFocus = { ...emptyFocusTecnico(), title: ex.title || "Nuovo Focus", exercises: [{ ...ex, id: uid("ex"), sourceExerciseId: ex.id }] };
    saveFocusTecnico(newFocus);
    showToast(`Nuovo Focus "${newFocus.title}" creato — rinominalo da Focus Tecnici se vuoi`);
  }

  return (
    <div>
      <SectionTitle
        eyebrow="Preparazione"
        title="Allenamenti"
        icon={Activity}
        action={
          subTab === "sessioni" && (
            <div className="flex gap-2 flex-wrap">
              <Button onClick={() => setShowAdd(true)}>
                <Plus className="w-4 h-4" /> Crea Nuovo Allenamento
              </Button>
              <SectionResetButton
                label="Azzera Allenamenti"
                confirmText="Eliminare tutte le sessioni di allenamento registrate?"
                onConfirm={() => {
                  updateSeason(() => ({ trainings: [] }));
                  showToast("Allenamenti azzerati");
                }}
              />
            </div>
          )
        }
      />

      <div className="flex gap-2 mb-5 flex-wrap">
        {[
          { id: "sessioni", label: "Sessioni" },
          { id: "focus", label: "Focus Tecnici" },
          { id: "esercizi", label: "Esercizi" },
          { id: "playbook", label: "PlayBook" },
        ].map((t) => (
          <button
            key={t.id}
            onClick={() => setSubTab(t.id)}
            className={`rounded-full px-4 py-2 text-xs font-semibold border transition-colors ${
              subTab === t.id ? "bg-emerald-500 text-slate-950 border-emerald-500" : "border-white/10 text-slate-400"
            }`}
          >
            {t.label}
          </button>
        ))}
      </div>

      {subTab === "playbook" ? (
        <PlayBookSection
          exercises={library.exercises || []}
          focusTecnici={focusTecnici}
          onSaveExercise={saveExercise}
          onAddToFocus={addExerciseToFocus}
          onCreateFocus={createFocusWithExercise}
          showToast={showToast}
        />
      ) : subTab === "focus" ? (
        <FocusTecniciSection
          focusTecnici={focusTecnici}
          onSave={saveFocusTecnico}
          onDelete={deleteFocusTecnico}
          exercises={library.exercises || []}
          onSaveExercise={saveExercise}
          onDeleteExercise={deleteExercise}
          showToast={showToast}
        />
      ) : subTab === "esercizi" ? (
        <ExercisesLibrarySection
          exercises={library.exercises || []}
          onSaveExercise={saveExercise}
          onDeleteExercise={deleteExercise}
          showToast={showToast}
        />
      ) : trainings.length === 0 ? (
        <EmptyState icon={Activity} text="Nessun allenamento registrato. Crea la prima sessione." />
      ) : (
        <div className="space-y-3">
          {trainings.map((t) => {
            const values = Object.values(t.attendance || {});
            const present = values.filter((v) => v === "Presente").length;
            const absent = values.filter((v) => v === "Assente").length;
            const pct = values.length ? Math.round((present / values.length) * 100) : 0;
            const linkedFocus = focusTecnici.find((f) => f.id === t.focusTecnicoId);
            return (
              <button key={t.id} onClick={() => setSelected(t)} className="w-full text-left">
                <Card className="p-4 hover:border-emerald-500/40 transition-colors">
                  <div className="flex items-center justify-between gap-4 flex-wrap">
                    <div className="flex items-center gap-3">
                      <div className="w-11 h-11 rounded-xl bg-sky-500/15 flex items-center justify-center shrink-0">
                        <Activity className="w-5 h-5 text-sky-400" />
                      </div>
                      <div>
                        <p className="text-sm font-bold text-slate-100">{formatDate(t.date)} · {t.time || "--:--"}</p>
                        <p className="text-xs text-slate-500">{linkedFocus?.title || t.focus || "Nessun focus indicato"}</p>
                      </div>
                    </div>
                    <div className="flex items-center gap-3">
                      <div className="text-right">
                        <p className="text-xs text-slate-500">Presenti {present} · Assenti {absent}</p>
                        <p className="text-sm font-bold text-emerald-400">{pct}%</p>
                      </div>
                      <ChevronRight className="w-4 h-4 text-slate-600" />
                    </div>
                  </div>
                </Card>
              </button>
            );
          })}
        </div>
      )}

      <Modal open={showAdd} onClose={() => setShowAdd(false)} title="Crea Nuovo Allenamento">
        <TrainingForm onSubmit={addTraining} onCancel={() => setShowAdd(false)} focusTecnici={focusTecnici} />
      </Modal>

      <Modal open={!!selected} onClose={() => setSelected(null)} title="Dettaglio Allenamento" wide>
        {selected && (
          <TrainingDetail
            training={trainings.find((t) => t.id === selected.id) || selected}
            players={players}
            focusTecnici={focusTecnici}
            onUpdate={(patch) => {
              updateTraining(selected.id, patch);
              setSelected((prev) => ({ ...prev, ...patch }));
            }}
            onDelete={() => deleteTraining(selected.id)}
          />
        )}
      </Modal>
    </div>
  );
}

/* ---------- Focus Tecnici (libreria esercizi) ---------- */

function emptyFocusTecnico() {
  return { id: uid("focus"), title: "", exercises: [{ id: uid("ex"), title: "", type: "Tecnica", time: "", goal: "", description: "" }] };
}

// Estrae il primo numero da una stringa tipo "10 min" o "15 minuti (3 serie)" per sommare le durate
function parseMinutes(timeStr) {
  if (!timeStr) return 0;
  const match = String(timeStr).match(/\d+/);
  return match ? parseInt(match[0], 10) : 0;
}

function totalFocusMinutes(ft) {
  return (ft.exercises || []).reduce((sum, ex) => sum + parseMinutes(ex.time), 0);
}

function emptyExercise() {
  return { id: uid("ex"), title: "", type: "Tecnica", category: "", time: "", goal: "", description: "", image: null };
}

function FocusTecniciSection({ focusTecnici, onSave, onDelete, exercises, onSaveExercise, onDeleteExercise, showToast }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [expandedFocus, setExpandedFocus] = useState(null);
  const [lightboxSrc, setLightboxSrc] = useState(null);

  const [showExerciseForm, setShowExerciseForm] = useState(false);
  const [editingExercise, setEditingExercise] = useState(null);
  const [confirmDeleteExId, setConfirmDeleteExId] = useState(null);

  return (
    <div>
      <div className="flex justify-end gap-2 mb-4 flex-wrap">
        <Button
          variant="secondary"
          onClick={() => {
            setEditingExercise(emptyExercise());
            setShowExerciseForm(true);
          }}
        >
          <Plus className="w-4 h-4" /> Crea Esercizio Singolo
        </Button>
        <Button
          onClick={() => {
            setEditing(emptyFocusTecnico());
            setShowForm(true);
          }}
        >
          <Plus className="w-4 h-4" /> Crea Focus Tecnico
        </Button>
      </div>

      <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide mb-2">Focus Tecnici (sessioni complete)</p>
      {focusTecnici.length === 0 ? (
        <EmptyState icon={Target} text="Nessun Focus Tecnico creato. Crea il primo per riutilizzarlo negli allenamenti." />
      ) : (
        <div className="grid sm:grid-cols-2 gap-4">
          {focusTecnici.map((ft) => (
            <Card key={ft.id} className="p-4">
              <div className="flex items-start justify-between gap-2 mb-2">
                <div>
                  <p className="text-sm font-bold text-slate-100">{ft.title || "Senza titolo"}</p>
                  <p className="text-[11px] text-emerald-400 font-medium mt-0.5">Durata totale: {totalFocusMinutes(ft)} min</p>
                </div>
                <div className="flex gap-1 shrink-0">
                  <button
                    onClick={() => setExpandedFocus(ft)}
                    className="rounded-lg p-1.5 hover:bg-white/10 text-slate-400"
                    title="Espandi (immagini grandi in sequenza)"
                  >
                    <Maximize2 className="w-3.5 h-3.5" />
                  </button>
                  <button
                    onClick={async () => {
                      const withImages = (ft.exercises || []).filter((ex) => ex.image);
                      if (withImages.length === 0) return showToast?.("Nessuna immagine presente in questo Focus", "error");
                      const summaryFile = dataUrlToFile(generateFocusSummaryImage(ft), "0-riepilogo-focus.png");
                      const exerciseFiles = withImages.map((ex, i) =>
                        dataUrlToFile(ex.image, `${(ex.title || `esercizio-${i + 1}`).replace(/[^a-z0-9]+/gi, "-")}.jpg`)
                      );
                      const files = [summaryFile, ...exerciseFiles];
                      const shared = await shareFilesNative(files, { title: ft.title, text: ft.title });
                      if (!shared) {
                        const baseName = (ft.title || "focus").replace(/[^a-z0-9]+/gi, "-");
                        const a0 = document.createElement("a");
                        a0.href = generateFocusSummaryImage(ft);
                        a0.download = `${baseName}-0-riepilogo.png`;
                        document.body.appendChild(a0);
                        a0.click();
                        document.body.removeChild(a0);
                        withImages.forEach((ex, i) => {
                          setTimeout(() => {
                            const a = document.createElement("a");
                            a.href = ex.image;
                            a.download = `${baseName}-${i + 1}.jpg`;
                            document.body.appendChild(a);
                            a.click();
                            document.body.removeChild(a);
                          }, (i + 1) * 350);
                        });
                        openWhatsAppFallback(`${ft.title || "Focus Tecnico"} — allega le immagini appena scaricate`);
                        showToast?.("Condivisione diretta non disponibile: immagini scaricate, allegale su WhatsApp manualmente.", "error");
                      }
                    }}
                    className="rounded-lg p-1.5 hover:bg-white/10 text-slate-400"
                    title="Condividi tutte le immagini (es. su WhatsApp)"
                  >
                    <Share2 className="w-3.5 h-3.5" />
                  </button>
                  <button onClick={() => downloadFocusSheet(ft)} className="rounded-lg p-1.5 hover:bg-white/10 text-slate-400" title="Scarica scheda completa">
                    <Download className="w-3.5 h-3.5" />
                  </button>
                  <button
                    onClick={() => {
                      setEditing(ft);
                      setShowForm(true);
                    }}
                    className="rounded-lg p-1.5 hover:bg-white/10 text-slate-400"
                  >
                    <Edit2 className="w-3.5 h-3.5" />
                  </button>
                  <button onClick={() => setConfirmDeleteId(ft.id)} className="rounded-lg p-1.5 hover:bg-rose-500/10 text-rose-400">
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                </div>
              </div>
              <div className="space-y-1.5">
                {(ft.exercises || []).map((ex, i) => (
                  <div key={ex.id || i} className="text-xs text-slate-400 flex items-start gap-2">
                    {ex.image && (
                      <button type="button" onClick={() => setLightboxSrc(ex.image)} className="shrink-0 mt-0.5">
                        <img src={ex.thumbnail || ex.image} alt={ex.title} className="w-9 h-9 object-cover rounded-lg" />
                      </button>
                    )}
                    <div className="min-w-0">
                      <div className="flex gap-1.5 flex-wrap items-center">
                        {ex.type && <Badge className={EXERCISE_TYPE_STYLES[ex.type] || EXERCISE_TYPE_STYLES.Tecnica}>{ex.type}</Badge>}
                        <span className="text-emerald-400 font-semibold shrink-0">{ex.title || "Esercizio"}</span>
                        <span className="text-slate-600">·</span>
                        <span className="shrink-0">{ex.time || "--"}</span>
                        {ex.station && <Badge className={STATION_COLORS[ex.station]}>Stazione {ex.station}</Badge>}
                      </div>
                      {ex.goal && <p className="text-[11px] text-slate-500 italic">Obiettivo: {ex.goal}</p>}
                    </div>
                  </div>
                ))}
              </div>
              {confirmDeleteId === ft.id && (
                <div className="mt-3 rounded-lg border border-rose-500/30 bg-rose-500/10 p-2.5 flex items-center justify-between gap-2">
                  <p className="text-xs text-rose-300">Eliminare?</p>
                  <div className="flex gap-1.5">
                    <Button variant="secondary" className="px-2 py-1 text-xs" onClick={() => setConfirmDeleteId(null)}>Annulla</Button>
                    <Button
                      variant="danger"
                      className="px-2 py-1 text-xs"
                      onClick={() => {
                        onDelete(ft.id);
                        setConfirmDeleteId(null);
                      }}
                    >
                      Elimina
                    </Button>
                  </div>
                </div>
              )}
            </Card>
          ))}
        </div>
      )}

      <Modal open={showForm} onClose={() => setShowForm(false)} title="Focus Tecnico" size="xl">
        {editing && (
          <FocusTecnicoForm
            initial={editing}
            standaloneExercises={exercises}
            onCreateExercise={onSaveExercise}
            onCancel={() => setShowForm(false)}
            onSubmit={(ft) => {
              onSave(ft);
              setShowForm(false);
            }}
          />
        )}
      </Modal>

      <Modal open={showExerciseForm} onClose={() => setShowExerciseForm(false)} title="Esercizio Singolo" wide>
        {editingExercise && (
          <ExerciseForm
            initial={editingExercise}
            onCancel={() => setShowExerciseForm(false)}
            onSubmit={(ex) => {
              onSaveExercise(ex);
              setShowExerciseForm(false);
            }}
          />
        )}
      </Modal>

      {expandedFocus && (
        <FocusDetailViewer focus={expandedFocus} onClose={() => setExpandedFocus(null)} showToast={showToast} />
      )}

      {lightboxSrc && <ImageLightbox src={lightboxSrc} alt="Esercizio" onClose={() => setLightboxSrc(null)} />}
    </div>
  );
}

/* ============================================================
   PLAYBOOK — catalogo di tutti gli esercizi, con vista "esplosa"
   simile a quella del Focus Tecnico, filtrabile per tipo/categoria
   ============================================================ */

function PlayBookSection({ exercises, focusTecnici, onSaveExercise, onAddToFocus, onCreateFocus, showToast }) {
  const config = useConfig();
  const [typeFilter, setTypeFilter] = useState("Tutti");
  const [categoryFilter, setCategoryFilter] = useState("Tutte");
  const [viewerIndex, setViewerIndex] = useState(null);

  const allTypesPresent = [
    "ND",
    ...Array.from(new Set([...config.exerciseTypes, ...(exercises || []).map((ex) => ex.type).filter(Boolean)])),
  ];
  const allCategoriesPresent = Array.from(
    new Set([...(config.categories || []), ...(exercises || []).map((ex) => ex.category).filter(Boolean)])
  );

  const matchesFilters = (ex) => {
    const matchesType = typeFilter === "Tutti" || (ex.type || "ND") === typeFilter;
    const matchesCategory = categoryFilter === "Tutte" || (categoryFilter === "ND" ? !ex.category : ex.category === categoryFilter);
    return matchesType && matchesCategory;
  };
  // Ordinato per tipologia (nell'ordine di allTypesPresent): così la navigazione
  // avanti/indietro nel visualizzatore esploso segue esattamente lo stesso ordine
  // mostrato nell'elenco raggruppato qui sotto, invece dell'ordine di inserimento.
  const filtered = allTypesPresent.flatMap((type) =>
    (exercises || []).filter((ex) => (ex.type || "ND") === type && matchesFilters(ex))
  );

  const grouped = allTypesPresent
    .map((type) => ({ type, items: filtered.filter((ex) => (ex.type || "ND") === type) }))
    .filter((g) => g.items.length > 0);

  return (
    <div>
      <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide mb-2">
        PlayBook — catalogo completo ({filtered.length} esercizi)
      </p>

      <div className="flex flex-wrap items-center gap-1.5 mb-2">
        {["Tutti", ...allTypesPresent].map((t) => (
          <button
            key={t}
            onClick={() => setTypeFilter(t)}
            className={`shrink-0 rounded-full px-3.5 py-1.5 text-[11px] font-medium border transition-colors ${
              typeFilter === t ? "bg-emerald-500 text-slate-950 border-emerald-500" : "border-white/10 text-slate-400 hover:text-slate-200"
            }`}
          >
            {t}
          </button>
        ))}
      </div>
      <div className="flex flex-wrap items-center gap-1.5 mb-5">
        {["Tutte", "ND", ...allCategoriesPresent].map((c) => (
          <button
            key={c}
            onClick={() => setCategoryFilter(c)}
            className={`shrink-0 rounded-full px-3.5 py-1.5 text-[11px] font-medium border transition-colors ${
              categoryFilter === c ? "bg-sky-500 text-slate-950 border-sky-500" : "border-white/10 text-slate-400 hover:text-slate-200"
            }`}
          >
            {c}
          </button>
        ))}
      </div>

      {grouped.length === 0 ? (
        <EmptyState icon={Target} text="Nessun esercizio trovato per questi filtri." />
      ) : (
        <div className="space-y-6">
          {grouped.map((g) => (
            <div key={g.type}>
              <div className="flex items-center gap-2 mb-2.5">
                <Badge className={EXERCISE_TYPE_STYLES[g.type] || NEUTRAL_BADGE}>{g.type}</Badge>
                <span className="text-xs text-slate-500">({g.items.length})</span>
              </div>
              <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-2">
                {g.items.map((ex) => (
                  <button
                    key={ex.id}
                    onClick={() => setViewerIndex(filtered.indexOf(ex))}
                    className="flex items-center gap-2.5 rounded-xl border border-white/10 p-2 text-left hover:bg-white/5"
                  >
                    {ex.image ? (
                      <img src={ex.thumbnail || ex.image} alt={ex.title} className="w-11 h-11 object-cover rounded-lg shrink-0" />
                    ) : (
                      <div className="w-11 h-11 rounded-lg bg-white/5 shrink-0 flex items-center justify-center">
                        <Target className="w-4 h-4 text-slate-600" />
                      </div>
                    )}
                    <div className="min-w-0">
                      <p className="text-sm font-medium text-slate-200 truncate">{ex.title}</p>
                      <p className="text-[11px] text-slate-500">{ex.time || "--"}{ex.category ? ` · ${ex.category}` : ""}</p>
                    </div>
                  </button>
                ))}
              </div>
            </div>
          ))}
        </div>
      )}

      {viewerIndex !== null && filtered[viewerIndex] && (
        <PlayBookViewer
          exercises={filtered}
          index={viewerIndex}
          onIndexChange={setViewerIndex}
          onClose={() => setViewerIndex(null)}
          onSaveExercise={onSaveExercise}
          focusTecnici={focusTecnici}
          onAddToFocus={onAddToFocus}
          onCreateFocus={onCreateFocus}
          showToast={showToast}
          allExercises={exercises || []}
          allTypesPresent={allTypesPresent}
        />
      )}
    </div>
  );
}

// Vista esplosa del PlayBook: un esercizio alla volta, immagine grande,
// navigazione avanti/indietro nell'insieme filtrato, con azioni rapide per
// modificare l'esercizio o aggiungerlo a un Focus Tecnico (esistente o nuovo).
function PlayBookViewer({ exercises, index, onIndexChange, onClose, onSaveExercise, focusTecnici, onAddToFocus, onCreateFocus, showToast, allExercises, allTypesPresent }) {
  const [showEdit, setShowEdit] = useState(false);
  const [showAddToFocus, setShowAddToFocus] = useState(false);
  const [showDownloadMenu, setShowDownloadMenu] = useState(false);
  const current = exercises[index];
  if (!current) return null;

  function goTo(delta) {
    const next = index + delta;
    if (next < 0 || next >= exercises.length) return;
    onIndexChange(next);
  }

  const currentType = current.type || "ND";
  const categoryItems = (allExercises || []).filter((ex) => (ex.type || "ND") === currentType);

  function handleDownloadCategory() {
    downloadCategorySheet(currentType, categoryItems);
    setShowDownloadMenu(false);
  }
  function handleDownloadPlayBook() {
    downloadPlayBookSheet(allExercises || [], allTypesPresent || [currentType]);
    setShowDownloadMenu(false);
  }
  function handleDownloadSingle() {
    downloadExerciseSheet(current);
    setShowDownloadMenu(false);
  }

  return (
    <div className="fixed inset-0 bg-black/95 flex flex-col" style={{ zIndex: 100 }}>
      <div className="flex items-center justify-between gap-2 px-4 py-2.5 border-b border-white/10 flex-wrap">
        <div className="min-w-0">
          <p className="text-sm font-bold text-white truncate">{current.type || "Esercizio"}</p>
          <p className="text-[11px] text-slate-400">Esercizio {index + 1} di {exercises.length}</p>
        </div>
        <div className="flex items-center gap-1.5 shrink-0">
          <button
            onClick={() =>
              shareOrFallback({
                dataUrl: current.image,
                filename: `${(current.title || "esercizio").replace(/[^a-z0-9]+/gi, "-")}.jpg`,
                label: current.title,
                showToast,
              })
            }
            disabled={!current.image}
            className="flex items-center gap-1.5 rounded-lg px-2.5 py-1.5 text-xs font-medium bg-white/10 hover:bg-white/20 text-white disabled:opacity-30"
            title="Condividi immagine"
          >
            <Share2 className="w-3.5 h-3.5" /> <span className="hidden sm:inline">Condividi</span>
          </button>
          <div className="relative">
            <button
              onClick={() => setShowDownloadMenu((v) => !v)}
              className="flex items-center gap-1.5 rounded-lg px-2.5 py-1.5 text-xs font-medium bg-white/10 hover:bg-white/20 text-white"
              title="Opzioni di download"
            >
              <Download className="w-3.5 h-3.5" /> <span className="hidden sm:inline">Scarica</span> <ChevronDown className="w-3 h-3" />
            </button>
            {showDownloadMenu && (
              <>
                <div className="fixed inset-0 z-10" onClick={() => setShowDownloadMenu(false)} />
                <div className="absolute right-0 top-full mt-1 w-56 rounded-xl border border-white/10 bg-slate-900 shadow-2xl overflow-hidden z-20">
                  <button onClick={handleDownloadSingle} className="w-full text-left px-3.5 py-2.5 text-xs text-slate-200 hover:bg-white/10">
                    Scarica Scheda <span className="block text-[10px] text-slate-500">Solo questo esercizio</span>
                  </button>
                  <button onClick={handleDownloadCategory} className="w-full text-left px-3.5 py-2.5 text-xs text-slate-200 hover:bg-white/10 border-t border-white/5">
                    Scarica Categoria <span className="block text-[10px] text-slate-500">Tutti gli esercizi "{currentType}" ({categoryItems.length})</span>
                  </button>
                  <button onClick={handleDownloadPlayBook} className="w-full text-left px-3.5 py-2.5 text-xs text-slate-200 hover:bg-white/10 border-t border-white/5">
                    Scarica PlayBook <span className="block text-[10px] text-slate-500">Tutti gli esercizi presenti ({(allExercises || []).length})</span>
                  </button>
                </div>
              </>
            )}
          </div>
          <button onClick={() => setShowAddToFocus(true)} className="flex items-center gap-1.5 rounded-lg px-2.5 py-1.5 text-xs font-medium bg-white/10 hover:bg-white/20 text-white" title="Aggiungi a un Focus Tecnico">
            <Plus className="w-3.5 h-3.5" /> <span className="hidden sm:inline">Aggiungi a Focus</span>
          </button>
          <button onClick={() => setShowEdit(true)} className="rounded-lg p-2 hover:bg-white/10 text-white" title="Modifica esercizio">
            <Edit2 className="w-4 h-4" />
          </button>
          <button onClick={onClose} className="rounded-lg p-2 hover:bg-white/10 text-white">
            <X className="w-5 h-5" />
          </button>
        </div>
      </div>

      <div className="flex-1 flex items-center justify-center relative px-4 min-h-0">
        {index > 0 && (
          <button onClick={() => goTo(-1)} className="absolute left-2 sm:left-4 rounded-full p-2 bg-white/10 hover:bg-white/20 text-white z-10">
            <ChevronLeft className="w-6 h-6" />
          </button>
        )}
        <div className="max-w-3xl w-full flex flex-col items-center gap-2.5 py-3 overflow-y-auto" style={{ maxHeight: "100%" }}>
          {current.image ? (
            <img src={current.image} alt={current.title} className="max-w-full object-contain rounded-lg" style={{ maxHeight: "52vh" }} />
          ) : (
            <div className="w-full aspect-video rounded-lg bg-white/5 flex items-center justify-center text-slate-500 text-sm">
              Nessuna immagine per questo esercizio
            </div>
          )}
          <div className="w-full text-left px-1">
            <div className="flex items-center gap-2 flex-wrap mb-1">
              {current.type && <Badge className={EXERCISE_TYPE_STYLES[current.type] || EXERCISE_TYPE_STYLES.Tecnica}>{current.type}</Badge>}
              {current.category && <Badge className="bg-sky-500/15 text-sky-400 border-sky-500/30">{current.category}</Badge>}
              <p className="text-base font-bold text-white">{current.title || "Esercizio"}</p>
              <span className="text-xs text-slate-400">· {current.time || "--"}</span>
            </div>
            {current.goal && <p className="text-sm text-emerald-400 italic mb-1">Obiettivo: {current.goal}</p>}
            {current.description && <p className="text-sm text-slate-300">{current.description}</p>}
          </div>
        </div>
        {index < exercises.length - 1 && (
          <button onClick={() => goTo(1)} className="absolute right-2 sm:right-4 rounded-full p-2 bg-white/10 hover:bg-white/20 text-white z-10">
            <ChevronRight className="w-6 h-6" />
          </button>
        )}
      </div>

      <Modal open={showEdit} onClose={() => setShowEdit(false)} title="Modifica Esercizio" wide>
        <ExerciseForm
          initial={current}
          onCancel={() => setShowEdit(false)}
          onSubmit={(updated) => {
            onSaveExercise(updated);
            setShowEdit(false);
          }}
        />
      </Modal>

      <Modal open={showAddToFocus} onClose={() => setShowAddToFocus(false)} title="Aggiungi a un Focus Tecnico">
        <div className="space-y-2">
          <button
            onClick={() => {
              onCreateFocus(current);
              setShowAddToFocus(false);
            }}
            className="w-full flex items-center gap-2 rounded-xl border border-emerald-500/30 bg-emerald-500/10 hover:bg-emerald-500/20 px-3 py-2.5 text-sm text-emerald-300 font-medium"
          >
            <Plus className="w-4 h-4" /> Crea nuovo Focus con questo esercizio
          </button>
          {(focusTecnici || []).length > 0 && (
            <>
              <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide pt-2">Oppure aggiungi a un Focus esistente</p>
              {focusTecnici.map((ft) => {
                const full = (ft.exercises || []).length >= 8;
                return (
                  <button
                    key={ft.id}
                    onClick={() => {
                      onAddToFocus(ft.id, current);
                      setShowAddToFocus(false);
                    }}
                    disabled={full}
                    className="w-full flex items-center justify-between gap-2 rounded-xl border border-white/10 hover:bg-white/5 px-3 py-2.5 text-sm text-slate-200 disabled:opacity-40 disabled:cursor-not-allowed"
                  >
                    <span className="truncate">{ft.title || "Senza titolo"}</span>
                    <span className="text-[11px] text-slate-500 shrink-0">{full ? "Pieno (8/8)" : `${(ft.exercises || []).length}/8`}</span>
                  </button>
                );
              })}
            </>
          )}
        </div>
      </Modal>
    </div>
  );
}

function ExerciseForm({ initial, onSubmit, onCancel }) {
  const config = useConfig();
  const [form, setForm] = useState(initial);
  const [imageError, setImageError] = useState("");
  const [showLightbox, setShowLightbox] = useState(false);
  const fileInputRef = React.useRef(null);

  async function handleImageChange(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    if (!file.type.startsWith("image/")) {
      setImageError("Seleziona un file immagine valido");
      return;
    }
    try {
      setImageError("");
      // Versione a piena risoluzione per ingrandimento/download/stampa, e miniatura compatta per l'elenco
      const [fullImage, thumbnail] = await Promise.all([
        resizeImageFile(file, 1600, 0.88),
        resizeImageFile(file, 220, 0.75),
      ]);
      setForm((f) => ({ ...f, image: fullImage, thumbnail }));
    } catch (err) {
      setImageError("Impossibile caricare l'immagine");
    }
    e.target.value = "";
  }

  return (
    <div className="grid sm:grid-cols-2 gap-x-6">
      <div>
        <Field label="Titolo esercizio">
          <input className={inputClass} value={form.title} onChange={(e) => setForm({ ...form, title: e.target.value })} placeholder="Es. Rondo 3v1" />
        </Field>
        <Field label="Tipologia">
          <select className={inputClass} value={form.type || "Tecnica"} onChange={(e) => setForm({ ...form, type: e.target.value })}>
            {config.exerciseTypes.map((t) => (
              <option key={t} value={t}>{t}</option>
            ))}
          </select>
        </Field>
        <Field label="Categoria">
          <select className={inputClass} value={form.category || ""} onChange={(e) => setForm({ ...form, category: e.target.value })}>
            <option value="">Nessuna categoria</option>
            {(config.categories || []).map((c) => (
              <option key={c} value={c}>{c}</option>
            ))}
          </select>
        </Field>
        <Field label="Tempo di esecuzione">
          <input className={inputClass} value={form.time} onChange={(e) => setForm({ ...form, time: e.target.value })} placeholder="Es. 10 minuti" />
        </Field>
      </div>

      <div>
        <Field label="Obiettivo">
          <input className={inputClass} value={form.goal || ""} onChange={(e) => setForm({ ...form, goal: e.target.value })} placeholder="Es. Migliorare il controllo orientato" />
        </Field>
        <Field label="Descrizione">
          <textarea rows={5} className={inputClass} value={form.description} onChange={(e) => setForm({ ...form, description: e.target.value })} placeholder="Es. Palleggi liberi" />
        </Field>
        <Field label="Immagine (facoltativa)">
          {form.image ? (
            <div>
              <button type="button" onClick={() => setShowLightbox(true)} className="block w-full">
                <img src={form.image} alt="Anteprima esercizio" className="w-full max-h-40 object-contain rounded-xl border border-white/10 bg-slate-950/40 mb-2" />
              </button>
              <div className="flex gap-2">
                <Button type="button" variant="secondary" onClick={() => fileInputRef.current?.click()}>Sostituisci</Button>
                <Button type="button" variant="ghost" onClick={() => setForm((f) => ({ ...f, image: null, thumbnail: null }))}>Rimuovi</Button>
              </div>
            </div>
          ) : (
            <Button type="button" variant="secondary" onClick={() => fileInputRef.current?.click()}>Carica immagine</Button>
          )}
          <input ref={fileInputRef} type="file" accept="image/*" className="hidden" onChange={handleImageChange} />
          {imageError && <p className="text-[11px] text-rose-400 mt-1">{imageError}</p>}
        </Field>

        <div className="flex flex-wrap justify-between gap-2 mt-4">
          <Button variant="secondary" onClick={() => downloadExerciseSheet(form)} title="Stampa / Scarica scheda esercizio">
            <Download className="w-4 h-4" /> Stampa
          </Button>
          <div className="flex gap-2">
            <Button variant="secondary" onClick={onCancel}>Annulla</Button>
            <Button onClick={() => form.title.trim() && onSubmit(form)}>
              <Save className="w-4 h-4" /> Salva Esercizio
            </Button>
          </div>
        </div>
      </div>

      {showLightbox && form.image && (
        <ImageLightbox src={form.image} alt={form.title || "Esercizio"} onClose={() => setShowLightbox(false)} />
      )}
    </div>
  );
}

function ExercisesLibrarySection({ exercises, onSaveExercise, onDeleteExercise, showToast }) {
  const config = useConfig();
  const [search, setSearch] = useState("");
  const [typeFilter, setTypeFilter] = useState("Tutti");
  const [categoryFilter, setCategoryFilter] = useState("Tutte");
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [confirmDeleteKey, setConfirmDeleteKey] = useState(null);
  const [lightboxSrc, setLightboxSrc] = useState(null);

  // Unica fonte di verità: tutti gli esercizi vivono qui come "Esercizio Singolo".
  // I Focus Tecnici possono solo referenziarli (vedi FocusTecnicoForm), non crearne
  // di nuovi al volo, per evitare duplicati come accadeva in passato.
  const combined = (exercises || []).map((ex) => ({ ...ex, _key: `standalone-${ex.id}` }));

  const filtered = combined.filter((ex) => {
    const q = search.toLowerCase();
    const matchesSearch =
      !q ||
      (ex.title || "").toLowerCase().includes(q) ||
      (ex.goal || "").toLowerCase().includes(q) ||
      (ex.description || "").toLowerCase().includes(q);
    const effectiveType = ex.type || "ND";
    const matchesType = typeFilter === "Tutti" || effectiveType === typeFilter;
    const matchesCategory = categoryFilter === "Tutte" || (categoryFilter === "ND" ? !ex.category : ex.category === categoryFilter);
    return matchesSearch && matchesType && matchesCategory;
  });

  // Unione tra le tipologie configurate e quelle eventualmente già usate nei dati
  // (per non nascondere esercizi con un tipo non più/non ancora presente in Configurazioni).
  // "ND" raggruppa gli esercizi senza tipologia assegnata.
  const allTypesPresent = [
    "ND",
    ...Array.from(new Set([...config.exerciseTypes, ...combined.map((ex) => ex.type).filter(Boolean)])),
  ];
  const allCategoriesPresent = Array.from(
    new Set([...(config.categories || []), ...combined.map((ex) => ex.category).filter(Boolean)])
  );
  const grouped = allTypesPresent
    .map((type) => ({
      type,
      items: filtered.filter((ex) => (ex.type || "ND") === type),
    }))
    .filter((g) => g.items.length > 0);

  function openEdit(ex) {
    setEditing(ex);
    setShowForm(true);
  }

  function handleSubmit(updatedEx) {
    onSaveExercise(updatedEx);
    setShowForm(false);
  }

  function handleDelete(ex) {
    onDeleteExercise(ex.id);
    setConfirmDeleteKey(null);
  }

  return (
    <div>
      <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
        <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide">
          Esercizi totali: <span className="text-emerald-400 font-bold">{combined.length}</span>
        </p>
      </div>

      <div className="flex flex-col sm:flex-row gap-3 mb-3">
        <div className="relative flex-1">
          <Search className="w-4 h-4 text-slate-500 absolute left-3 top-1/2 -translate-y-1/2" />
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Cerca esercizio per titolo, obiettivo o descrizione..."
            className={inputClass + " pl-9"}
          />
        </div>
        <Button
          onClick={() => {
            setEditing(emptyExercise());
            setEditingSource(null);
            setShowForm(true);
          }}
        >
          <Plus className="w-4 h-4" /> Crea Esercizio Singolo
        </Button>
      </div>

      <div className="flex items-center gap-2 mb-2 overflow-x-auto pb-1">
        {["Tutti", ...allTypesPresent].map((t) => (
          <button
            key={t}
            onClick={() => setTypeFilter(t)}
            className={`shrink-0 rounded-full px-3.5 py-1.5 text-[11px] font-medium border transition-colors ${
              typeFilter === t
                ? "bg-emerald-500 text-slate-950 border-emerald-500"
                : "border-white/10 text-slate-400 hover:text-slate-200"
            }`}
          >
            {t}
          </button>
        ))}
      </div>

      <div className="flex items-center gap-2 mb-5 overflow-x-auto pb-1">
        {["Tutte", "ND", ...allCategoriesPresent].map((c) => (
          <button
            key={c}
            onClick={() => setCategoryFilter(c)}
            className={`shrink-0 rounded-full px-3.5 py-1.5 text-[11px] font-medium border transition-colors ${
              categoryFilter === c
                ? "bg-sky-500 text-slate-950 border-sky-500"
                : "border-white/10 text-slate-400 hover:text-slate-200"
            }`}
          >
            {c}
          </button>
        ))}
      </div>

      {grouped.length === 0 ? (
        <EmptyState icon={Target} text={search ? "Nessun esercizio trovato per questa ricerca." : "Nessun esercizio creato ancora."} />
      ) : (
        <div className="space-y-6">
          {grouped.map((g) => (
            <div key={g.type}>
              <div className="flex items-center gap-2 mb-2.5">
                <Badge className={EXERCISE_TYPE_STYLES[g.type] || NEUTRAL_BADGE}>{g.type}</Badge>
                <span className="text-xs text-slate-500">({g.items.length})</span>
              </div>
              <div className="space-y-2">
                {g.items.map((ex) => (
                  <Card key={ex._key} className="p-2.5">
                    <div className="flex items-center gap-3">
                      {ex.image && (
                        <button type="button" onClick={() => setLightboxSrc(ex.image)} className="shrink-0">
                          <img src={ex.thumbnail || ex.image} alt={ex.title} className="w-10 h-10 object-cover rounded-lg" />
                        </button>
                      )}
                      <div className="flex-1 min-w-0 flex items-center justify-between gap-2 flex-wrap">
                        <div className="flex items-center gap-2 flex-wrap min-w-0">
                          <span className="text-sm font-semibold text-slate-100 whitespace-nowrap">{ex.title || "Senza titolo"}</span>
                          <span className="text-[11px] text-slate-500 whitespace-nowrap">{ex.time || "--"}</span>
                          {ex.goal && <span className="text-[11px] text-slate-500 italic">Obiettivo: {ex.goal}</span>}
                          {ex.category && (
                            <Badge className="bg-sky-500/15 text-sky-400 border-sky-500/30">{ex.category}</Badge>
                          )}
                          <Badge className="bg-white/5 text-slate-400 border-white/10">
                            Esercizio singolo
                          </Badge>
                        </div>
                        <div className="flex gap-1 shrink-0">
                          {ex.image && (
                            <button
                              onClick={() => shareOrFallback({ dataUrl: ex.image, filename: `${(ex.title || "esercizio").replace(/[^a-z0-9]+/gi, "-")}.jpg`, label: ex.title, showToast })}
                              className="rounded-lg p-1.5 hover:bg-white/10 text-slate-400"
                              title="Condividi immagine (es. su WhatsApp)"
                            >
                              <Share2 className="w-3.5 h-3.5" />
                            </button>
                          )}
                          <button onClick={() => downloadExerciseSheet(ex)} className="rounded-lg p-1.5 hover:bg-white/10 text-slate-400" title="Stampa / Scarica scheda esercizio">
                            <Download className="w-3.5 h-3.5" />
                          </button>
                          <button onClick={() => openEdit(ex)} className="rounded-lg p-1.5 hover:bg-white/10 text-slate-400">
                            <Edit2 className="w-3.5 h-3.5" />
                          </button>
                          <button onClick={() => setConfirmDeleteKey(ex._key)} className="rounded-lg p-1.5 hover:bg-rose-500/10 text-rose-400">
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </div>
                    </div>
                    {confirmDeleteKey === ex._key && (
                      <div className="mt-2 rounded-lg border border-rose-500/30 bg-rose-500/10 p-2 flex items-center justify-between gap-2">
                        <p className="text-[11px] text-rose-300">Eliminare?</p>
                        <div className="flex gap-1.5">
                          <Button variant="secondary" className="px-2 py-1 text-xs" onClick={() => setConfirmDeleteKey(null)}>Annulla</Button>
                          <Button variant="danger" className="px-2 py-1 text-xs" onClick={() => handleDelete(ex)}>
                            Elimina
                          </Button>
                        </div>
                      </div>
                    )}
                  </Card>
                ))}
              </div>
            </div>
          ))}
        </div>
      )}

      <Modal open={showForm} onClose={() => setShowForm(false)} title="Esercizio Singolo" wide>
        {editing && (
          <ExerciseForm
            initial={editing}
            onCancel={() => setShowForm(false)}
            onSubmit={handleSubmit}
          />
        )}
      </Modal>

      {lightboxSrc && <ImageLightbox src={lightboxSrc} alt="Esercizio" onClose={() => setLightboxSrc(null)} />}
    </div>
  );
}

const STATION_LETTERS = ["A", "B", "C", "D"];
const STATION_COLORS = {
  A: "bg-amber-500/20 text-amber-300 border-amber-500/40",
  B: "bg-purple-500/20 text-purple-300 border-purple-500/40",
  C: "bg-pink-500/20 text-pink-300 border-pink-500/40",
  D: "bg-cyan-500/20 text-cyan-300 border-cyan-500/40",
};
const MAX_FOCUS_EXERCISES = 8;

function FocusTecnicoForm({ initial, onSubmit, onCancel, standaloneExercises, onCreateExercise }) {
  const config = useConfig();
  const [form, setForm] = useState(initial);
  const [search, setSearch] = useState("");
  const [typeFilter, setTypeFilter] = useState("Tutti");
  const [showCreateExercise, setShowCreateExercise] = useState(false);
  const [lightboxSrc, setLightboxSrc] = useState(null);

  const selectedSourceIds = new Set(form.exercises.map((ex) => ex.sourceExerciseId).filter(Boolean));
  const totalMinutes = form.exercises.reduce((sum, ex) => sum + parseMinutes(ex.time), 0);

  const filteredLibrary = (standaloneExercises || [])
    .filter((ex) => ex.title)
    .filter((ex) => {
      const q = search.trim().toLowerCase();
      const matchesSearch = !q || ex.title.toLowerCase().includes(q) || (ex.goal || "").toLowerCase().includes(q);
      const matchesType = typeFilter === "Tutti" || (ex.type || "Tecnica") === typeFilter;
      return matchesSearch && matchesType;
    })
    .sort((a, b) => (a.type || "Tecnica").localeCompare(b.type || "Tecnica", "it") || a.title.localeCompare(b.title, "it"));

  function addExistingExercise(source) {
    if (form.exercises.length >= MAX_FOCUS_EXERCISES || selectedSourceIds.has(source.id)) return;
    setForm({ ...form, exercises: [...form.exercises, { ...source, id: uid("ex"), sourceExerciseId: source.id }] });
  }

  function removeExercise(idx) {
    setForm({ ...form, exercises: form.exercises.filter((_, i) => i !== idx) });
  }

  function moveExercise(idx, direction) {
    const newIdx = idx + direction;
    if (newIdx < 0 || newIdx >= form.exercises.length) return;
    const exercises = [...form.exercises];
    [exercises[idx], exercises[newIdx]] = [exercises[newIdx], exercises[idx]];
    setForm({ ...form, exercises });
  }

  // Cicla l'assegnazione "stazione" di un esercizio: nessuna -> A -> B -> C -> D -> nessuna.
  // Serve a segnalare quando 2+ esercizi vengono svolti in parallelo su stazioni diverse
  // del campo, invece che in sequenza uno dopo l'altro.
  function cycleStation(idx) {
    const exercises = [...form.exercises];
    const current = exercises[idx].station;
    const currentPos = STATION_LETTERS.indexOf(current);
    const next = currentPos === -1 ? STATION_LETTERS[0] : currentPos === STATION_LETTERS.length - 1 ? null : STATION_LETTERS[currentPos + 1];
    exercises[idx] = { ...exercises[idx], station: next };
    setForm({ ...form, exercises });
  }

  function handleNewExerciseCreated(newEx) {
    onCreateExercise(newEx); // salva nella libreria Esercizi Singoli
    if (form.exercises.length < MAX_FOCUS_EXERCISES) {
      setForm((f) => ({ ...f, exercises: [...f.exercises, { ...newEx, id: uid("ex"), sourceExerciseId: newEx.id }] }));
    }
    setShowCreateExercise(false);
  }

  return (
    <div>
      <div className="flex items-center justify-between gap-2 mb-4">
        <h4 className="text-sm font-bold text-slate-100">Titolo dell'allenamento tipo</h4>
        <div className="flex gap-2 shrink-0">
          <Button variant="secondary" onClick={onCancel}>Annulla</Button>
          <Button onClick={() => form.title.trim() && onSubmit(form)}>
            <Save className="w-4 h-4" /> Salva Focus Tecnico
          </Button>
        </div>
      </div>
      <div className="grid lg:grid-cols-2 gap-5">
        {/* COLONNA SINISTRA: titolo e riepilogo esercizi selezionati */}
        <div>
          <input
            className={inputClass + " mb-4"}
            value={form.title}
            onChange={(e) => setForm({ ...form, title: e.target.value })}
            placeholder="Es. Possesso palla e transizioni"
          />

          {form.exercises.length > 0 ? (
            <div>
              <div className="flex items-center justify-between mb-2 flex-wrap gap-1">
                <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide">
                  Esercizi selezionati ({form.exercises.length}/{MAX_FOCUS_EXERCISES})
                </p>
                <p className="text-xs font-semibold text-emerald-400">Durata totale: {totalMinutes} min</p>
              </div>
              <div className="space-y-2">
                {form.exercises.map((ex, i) => (
                  <div key={ex.id || i} className="rounded-xl border border-white/10 p-2.5 flex items-center gap-3 flex-wrap">
                    <span className="text-xs font-bold text-slate-500 w-4 shrink-0 text-center">{i + 1}</span>
                    {ex.image && (
                      <button type="button" onClick={() => setLightboxSrc(ex.image)} className="shrink-0">
                        <img src={ex.thumbnail || ex.image} alt={ex.title} className="w-10 h-10 object-cover rounded-lg" />
                      </button>
                    )}
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-1.5 flex-wrap">
                        {ex.type && <Badge className={EXERCISE_TYPE_STYLES[ex.type] || EXERCISE_TYPE_STYLES.Tecnica}>{ex.type}</Badge>}
                        <span className="text-sm font-semibold text-slate-100">{ex.title}</span>
                        <span className="text-[11px] text-slate-500">· {ex.time || "--"}</span>
                      </div>
                      {ex.goal && <p className="text-[11px] text-slate-500 italic truncate">Obiettivo: {ex.goal}</p>}
                    </div>
                    <button
                      onClick={() => cycleStation(i)}
                      className={`shrink-0 rounded-lg border px-2.5 py-1.5 text-[11px] font-semibold transition-colors ${
                        ex.station ? STATION_COLORS[ex.station] : "border-white/10 text-slate-500 hover:text-slate-300"
                      }`}
                      title="Segna come stazione in parallelo con altri esercizi (es. più stazioni contemporanee sul campo)"
                    >
                      {ex.station ? `Stazione ${ex.station}` : "Stazioni"}
                    </button>
                    <div className="flex items-center gap-0.5 shrink-0">
                      <button onClick={() => moveExercise(i, -1)} disabled={i === 0} className="text-slate-400 hover:text-slate-200 p-1 disabled:opacity-20" title="Sposta su">
                        <ChevronUp className="w-3.5 h-3.5" />
                      </button>
                      <button onClick={() => moveExercise(i, 1)} disabled={i === form.exercises.length - 1} className="text-slate-400 hover:text-slate-200 p-1 disabled:opacity-20" title="Sposta giù">
                        <ChevronDown className="w-3.5 h-3.5" />
                      </button>
                      <button onClick={() => removeExercise(i)} className="text-rose-400 p-1" title="Rimuovi dal Focus">
                        <X className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          ) : (
            <EmptyState icon={Target} text="Nessun esercizio ancora selezionato: scegli dalla libreria qui a fianco." />
          )}
        </div>

        {/* COLONNA DESTRA: libreria esercizi da cui scegliere, su sfondo scuro per differenziarla */}
        <div className="rounded-2xl border border-white/10 bg-slate-950/70 p-4">
          <div className="flex items-center justify-between mb-2 flex-wrap gap-2">
            <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide">Aggiungi dalla libreria Esercizi</p>
            <button
              onClick={() => setShowCreateExercise(true)}
              className="text-emerald-400 text-xs font-medium flex items-center gap-1"
            >
              <Plus className="w-3.5 h-3.5" /> Non lo trovi? Crea nuovo esercizio
            </button>
          </div>

          <div className="relative mb-2">
            <Search className="w-4 h-4 text-slate-500 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Cerca esercizio per titolo o obiettivo..."
              className={inputClass + " pl-9"}
            />
          </div>
          <div className="flex flex-wrap items-center gap-1.5 mb-3">
            {["Tutti", ...config.exerciseTypes].map((t) => (
              <button
                key={t}
                onClick={() => setTypeFilter(t)}
                className={`shrink-0 rounded-full px-3 py-1 text-[11px] font-medium border transition-colors ${
                  typeFilter === t ? "bg-emerald-500 text-slate-950 border-emerald-500" : "border-white/10 text-slate-400 hover:text-slate-200"
                }`}
              >
                {t}
              </button>
            ))}
          </div>

          <div className="rounded-xl border border-white/10 bg-slate-900/60 max-h-[28rem] overflow-y-auto divide-y divide-white/5">
            {filteredLibrary.length === 0 ? (
              <p className="text-xs text-slate-500 text-center py-6">
                {(standaloneExercises || []).length === 0
                  ? "Nessun esercizio in libreria: creane uno con il pulsante qui sopra."
                  : "Nessun esercizio trovato per questa ricerca."}
              </p>
            ) : (
              filteredLibrary.map((ex) => {
                const alreadyAdded = selectedSourceIds.has(ex.id);
                return (
                  <button
                    key={ex.id}
                    type="button"
                    onClick={() => addExistingExercise(ex)}
                    disabled={alreadyAdded || form.exercises.length >= MAX_FOCUS_EXERCISES}
                    className="w-full flex items-center gap-3 px-3 py-2 text-left hover:bg-white/5 disabled:opacity-40 disabled:cursor-not-allowed"
                  >
                    {ex.image ? (
                      <img src={ex.thumbnail || ex.image} alt={ex.title} className="w-9 h-9 object-cover rounded-lg shrink-0" />
                    ) : (
                      <div className="w-9 h-9 rounded-lg bg-white/5 shrink-0" />
                    )}
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-1.5 flex-wrap">
                        {ex.type && <Badge className={EXERCISE_TYPE_STYLES[ex.type] || EXERCISE_TYPE_STYLES.Tecnica}>{ex.type}</Badge>}
                        <span className="text-sm text-slate-200 truncate">{ex.title}</span>
                      </div>
                    </div>
                    <span className="text-[11px] text-slate-500 shrink-0">{alreadyAdded ? "Già aggiunto" : ex.time || "--"}</span>
                  </button>
                );
              })
            )}
          </div>
        </div>
      </div>

      <Modal open={showCreateExercise} onClose={() => setShowCreateExercise(false)} title="Nuovo Esercizio Singolo" wide>
        <ExerciseForm initial={emptyExercise()} onCancel={() => setShowCreateExercise(false)} onSubmit={handleNewExerciseCreated} />
      </Modal>

      {lightboxSrc && <ImageLightbox src={lightboxSrc} alt="Esercizio" onClose={() => setLightboxSrc(null)} />}
    </div>
  );
}

function TrainingForm({ onSubmit, onCancel, focusTecnici, initial }) {
  const isEditing = !!initial;
  const [form, setForm] = useState(
    initial || { date: todayISO(), time: "17:00", focusTecnicoId: "", focus: "" }
  );
  return (
    <div>
      <Field label="Data">
        <input type="date" className={inputClass} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} />
      </Field>
      <Field label="Ora">
        <input type="time" className={inputClass} value={form.time} onChange={(e) => setForm({ ...form, time: e.target.value })} />
      </Field>
      <Field label="Focus Tecnico">
        <select
          className={inputClass}
          value={form.focusTecnicoId}
          onChange={(e) => setForm({ ...form, focusTecnicoId: e.target.value })}
        >
          <option value="">— Nessuno / nota libera —</option>
          {(focusTecnici || []).map((ft) => (
            <option key={ft.id} value={ft.id}>{ft.title}</option>
          ))}
        </select>
      </Field>
      {!form.focusTecnicoId && (
        <Field label="Nota libera (facoltativa)">
          <input className={inputClass} value={form.focus} onChange={(e) => setForm({ ...form, focus: e.target.value })} placeholder="Es. Possesso palla e transizioni" />
        </Field>
      )}
      <div className="flex justify-end gap-2 mt-4">
        <Button variant="secondary" onClick={onCancel}>Annulla</Button>
        <Button onClick={() => onSubmit(form)}>
          <Save className="w-4 h-4" /> {isEditing ? "Salva Modifiche" : "Crea Allenamento"}
        </Button>
      </div>
      {!isEditing && (
        <p className="text-[11px] text-slate-500 mt-3">Tutti i giocatori della rosa verranno impostati come "Presente" di default: potrai correggerli nel dettaglio.</p>
      )}
    </div>
  );
}

function TrainingDetail({ training, players, focusTecnici, onUpdate, onDelete }) {
  const attendance = training.attendance || {};
  const values = Object.values(attendance);
  const present = values.filter((v) => v === "Presente").length;
  const absent = values.filter((v) => v === "Assente").length;
  const justified = values.filter((v) => v === "Giustificato").length;
  const injured = values.filter((v) => v === "Infortunato").length;
  const pct = values.length ? Math.round((present / values.length) * 100) : 0;
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [expandedStatus, setExpandedStatus] = useState(null);
  const [editing, setEditing] = useState(false);
  const [exerciseLightboxSrc, setExerciseLightboxSrc] = useState(null);

  const linkedFocus = (focusTecnici || []).find((f) => f.id === training.focusTecnicoId);

  function setStatus(playerId, status) {
    onUpdate({ attendance: { ...attendance, [playerId]: status } });
  }

  function playersByStatus(status) {
    return players.filter((p) => attendance[p.id] === status);
  }

  const summary = [
    { status: "Presente", count: present, color: "text-emerald-400" },
    { status: "Assente", count: absent, color: "text-rose-400" },
    { status: "Giustificato", count: justified, color: "text-amber-400" },
    { status: "Infortunato", count: injured, color: "text-slate-400" },
  ];

  function downloadTrainingSheet() {
    let body = `<h1>Scheda Allenamento</h1><p>${formatDate(training.date)} · ${training.time}</p>`;
    body += `<p><strong>Presenti:</strong> ${present} · <strong>Assenti:</strong> ${absent} · <strong>Giustificati:</strong> ${justified} · <strong>Infortunati:</strong> ${injured}</p>`;
    if (linkedFocus) {
      body += `<h2>${linkedFocus.title} — Durata totale: ${totalFocusMinutes(linkedFocus)} min</h2><ol>`;
      (linkedFocus.exercises || []).forEach((ex, i) => {
        body += `<li><span class="badge">${ex.type || "Tecnica"}</span><strong>${ex.title || `Esercizio ${i + 1}`}</strong> — ${ex.time || "--"}`;
        if (ex.goal) body += `<div>Obiettivo: ${ex.goal}</div>`;
        if (ex.description) body += `<div>${ex.description}</div>`;
        body += `</li>`;
      });
      body += `</ol>`;
    } else {
      body += `<p>${training.focus || "Nessun focus indicato"}</p>`;
    }
    downloadPrintableHTML(`allenamento-${training.date}.html`, "Scheda Allenamento", body);
  }

  if (editing) {
    return (
      <TrainingForm
        initial={{ date: training.date, time: training.time, focusTecnicoId: training.focusTecnicoId || "", focus: training.focus || "" }}
        focusTecnici={focusTecnici}
        onCancel={() => setEditing(false)}
        onSubmit={(patch) => {
          onUpdate(patch);
          setEditing(false);
        }}
      />
    );
  }

  return (
    <div>
      <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
        <div>
          <p className="text-sm font-bold text-slate-100">{formatDate(training.date)} · {training.time}</p>
          <p className="text-xs text-slate-500">{linkedFocus?.title || training.focus || "Nessun focus indicato"}</p>
        </div>
        <div className="flex items-center gap-3">
          <Badge className="bg-emerald-500/15 text-emerald-400 border-emerald-500/30 text-sm">
            Presenza {pct}%
          </Badge>
          <button onClick={downloadTrainingSheet} className="rounded-lg p-2 hover:bg-white/10 text-slate-400" title="Scarica scheda allenamento (HTML stampabile)">
            <Download className="w-4 h-4" />
          </button>
          <button onClick={() => setEditing(true)} className="rounded-lg p-2 hover:bg-white/10 text-slate-400" title="Modifica allenamento">
            <Edit2 className="w-4 h-4" />
          </button>
          <button onClick={() => setConfirmDelete(true)} className="rounded-lg p-2 hover:bg-rose-500/10 text-rose-400" title="Elimina">
            <Trash2 className="w-4 h-4" />
          </button>
        </div>
      </div>

      {confirmDelete && (
        <div className="rounded-xl border border-rose-500/30 bg-rose-500/10 p-3 mb-4 flex items-center justify-between gap-2">
          <p className="text-sm text-rose-300">Eliminare questo allenamento?</p>
          <div className="flex gap-2 shrink-0">
            <Button variant="secondary" onClick={() => setConfirmDelete(false)}>Annulla</Button>
            <Button variant="danger" onClick={onDelete}>Elimina</Button>
          </div>
        </div>
      )}

      {/* RIEPILOGO PRESENZE */}
      <div className="grid grid-cols-4 gap-2 mb-4">
        {summary.map((s) => (
          <button
            key={s.status}
            onClick={() => setExpandedStatus(expandedStatus === s.status ? null : s.status)}
            className={`rounded-xl border p-2.5 text-center transition-colors ${
              expandedStatus === s.status ? "border-emerald-500/40 bg-white/5" : "border-white/10"
            }`}
          >
            <p className={`text-lg font-extrabold ${s.color}`}>{s.count}</p>
            <p className="text-[10px] text-slate-500">{s.status}</p>
          </button>
        ))}
      </div>

      {expandedStatus && (
        <div className="rounded-xl border border-white/10 p-3 mb-4">
          <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide mb-2">{expandedStatus}</p>
          {playersByStatus(expandedStatus).length === 0 ? (
            <p className="text-xs text-slate-500">Nessun giocatore in questo stato.</p>
          ) : (
            <div className="flex flex-wrap gap-2">
              {playersByStatus(expandedStatus).map((p) => (
                <span key={p.id} className="inline-flex items-center gap-1.5 rounded-full bg-white/5 pl-1 pr-2.5 py-1 text-xs text-slate-300">
                  <img src={playerAvatar(p)} className="w-5 h-5 rounded-full object-cover" />
                  {p.name}
                </span>
              ))}
            </div>
          )}
        </div>
      )}

      {/* ESERCIZI PREVISTI (dal Focus Tecnico collegato) */}
      {linkedFocus && (linkedFocus.exercises || []).length > 0 && (
        <div className="mb-5">
          <div className="flex items-center justify-between mb-2">
            <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide">Esercizi previsti</p>
            <span className="text-[11px] text-emerald-400 font-medium">Durata totale: {totalFocusMinutes(linkedFocus)} min</span>
          </div>
          <div className="space-y-2">
            {linkedFocus.exercises.map((ex, i) => (
              <div key={ex.id || i} className="rounded-xl border border-white/10 p-3 flex items-start gap-3">
                {ex.image && (
                  <button type="button" onClick={() => setExerciseLightboxSrc(ex.image)} className="shrink-0">
                    <img src={ex.thumbnail || ex.image} alt={ex.title} className="w-12 h-12 object-cover rounded-lg" />
                  </button>
                )}
                <div className="min-w-0 flex-1">
                  <div className="flex items-center justify-between gap-2 flex-wrap">
                    <div className="flex items-center gap-2 flex-wrap">
                      {ex.type && <Badge className={EXERCISE_TYPE_STYLES[ex.type] || EXERCISE_TYPE_STYLES.Tecnica}>{ex.type}</Badge>}
                      <p className="text-sm font-semibold text-slate-200">{ex.title || `Esercizio ${i + 1}`}</p>
                    </div>
                    {ex.time && <Badge className="bg-sky-500/15 text-sky-400 border-sky-500/30 shrink-0">{ex.time}</Badge>}
                  </div>
                  {ex.goal && <p className="text-xs text-emerald-400/90 mt-1 italic">Obiettivo: {ex.goal}</p>}
                  {ex.description && <p className="text-xs text-slate-500 mt-1">{ex.description}</p>}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide mb-2">Presenze giocatori</p>
      {players.length === 0 ? (
        <EmptyState icon={Users} text="Aggiungi giocatori alla rosa per registrare le presenze." />
      ) : (
        <div className="space-y-2 overflow-y-auto pr-1" style={{ maxHeight: "45vh" }}>
          {players.map((p) => (
            <div key={p.id} className="flex items-center justify-between gap-3 rounded-xl border border-white/10 p-2.5">
              <div className="flex items-center gap-2.5 min-w-0">
                <img src={playerAvatar(p)} alt={p.name} className="w-9 h-9 rounded-full bg-slate-800 object-cover" />
                <div className="min-w-0">
                  <p className="text-sm font-medium text-slate-200 truncate">{p.name}</p>
                  <p className="text-[11px] text-slate-500">{p.role}</p>
                </div>
              </div>
              <div className="flex gap-1 flex-wrap justify-end">
                {ATTENDANCE_STATUS.map((st) => (
                  <button
                    key={st}
                    onClick={() => setStatus(p.id, st)}
                    className={`rounded-full border px-2.5 py-1 text-[11px] font-medium transition-colors ${
                      attendance[p.id] === st ? ATTENDANCE_COLORS[st] : "border-white/10 text-slate-600 hover:text-slate-300"
                    }`}
                  >
                    {st}
                  </button>
                ))}
              </div>
            </div>
          ))}
        </div>
      )}

      {exerciseLightboxSrc && <ImageLightbox src={exerciseLightboxSrc} alt="Esercizio" onClose={() => setExerciseLightboxSrc(null)} />}
    </div>
  );
}

/* ============================================================
   SEZIONE PARTITE
   ============================================================ */

function MatchesSection({ season, updateSeason, showToast }) {
  const matches = [...(season.matches || [])].sort((a, b) => new Date(b.date) - new Date(a.date));
  const players = season.players || [];
  const [filter, setFilter] = useState("Programmata");
  const [typeFilter, setTypeFilter] = useState("Tutte");
  const [showAdd, setShowAdd] = useState(false);
  const [selected, setSelected] = useState(null);

  const filtered = matches
    .filter((m) => m.status === filter)
    .filter((m) => filter !== "Disputata" || typeFilter === "Tutte" || m.matchType === typeFilter);

  function addMatch(data) {
    updateSeason((s) => ({
      matches: [
        ...(s.matches || []),
        {
          ...data,
          id: uid("match"),
          status: "Programmata",
          result: null,
          scorers: [],
          assists: [],
          yellowCards: [],
          redCards: [],
          coachNotes: "",
        },
      ],
    }));
    setShowAdd(false);
    showToast("Partita programmata");
  }

  function updateMatch(id, patch) {
    updateSeason((s) => ({
      matches: (s.matches || []).map((m) => (m.id === id ? { ...m, ...patch } : m)),
    }));
  }

  function deleteMatch(id) {
    updateSeason((s) => ({ matches: (s.matches || []).filter((m) => m.id !== id) }));
    setSelected(null);
  }

  return (
    <div>
      <SectionTitle
        eyebrow="Risultati"
        title="Partite"
        icon={Trophy}
        action={
          <div className="flex gap-2 flex-wrap">
            <Button onClick={() => setShowAdd(true)}>
              <Plus className="w-4 h-4" /> Programma Partita
            </Button>
            <SectionResetButton
              label="Azzera Partite"
              confirmText="Eliminare tutte le partite, programmate e disputate?"
              onConfirm={() => {
                updateSeason(() => ({ matches: [] }));
                showToast("Partite azzerate");
              }}
            />
          </div>
        }
      />

      <div className="flex gap-2 mb-3">
        {["Programmata", "Disputata"].map((f) => (
          <button
            key={f}
            onClick={() => setFilter(f)}
            className={`rounded-full px-4 py-2 text-xs font-semibold border transition-colors ${
              filter === f ? "bg-emerald-500 text-slate-950 border-emerald-500" : "border-white/10 text-slate-400"
            }`}
          >
            {f === "Programmata" ? "Programmate" : "Disputate"}
          </button>
        ))}
      </div>

      {filter === "Disputata" && (
        <div className="flex gap-2 mb-5 flex-wrap">
          {["Tutte", ...MATCH_TYPES].map((t) => (
            <button
              key={t}
              onClick={() => setTypeFilter(t)}
              className={`rounded-full px-3.5 py-1.5 text-[11px] font-medium border transition-colors ${
                typeFilter === t ? "bg-white/10 text-slate-100 border-white/30" : "border-white/10 text-slate-500"
              }`}
            >
              {t}
            </button>
          ))}
        </div>
      )}
      {filter !== "Disputata" && <div className="mb-5" />}

      {filtered.length === 0 ? (
        <EmptyState icon={Trophy} text={`Nessuna partita ${filter === "Programmata" ? "programmata" : "disputata"}.`} />
      ) : (
        <div className="space-y-3">
          {filtered.map((m) => {
            const TypeIcon = MATCH_TYPE_ICONS[m.matchType] || Trophy;
            return (
              <button key={m.id} onClick={() => setSelected(m)} className="w-full text-left">
                <Card
                  className="p-4 hover:border-emerald-500/40 transition-colors"
                  style={{ borderLeft: `4px solid ${MATCH_TYPE_BORDER[m.matchType] || "transparent"}` }}
                >
                  <div className="flex items-center justify-between gap-4 flex-wrap">
                    <div className="flex items-center gap-3">
                      <div className={`w-11 h-11 rounded-xl flex items-center justify-center shrink-0 ${m.homeAway === "Casa" ? "bg-emerald-500/15" : "bg-amber-500/15"}`}>
                        <Flag className={`w-5 h-5 ${m.homeAway === "Casa" ? "text-emerald-400" : "text-amber-400"}`} />
                      </div>
                      <div>
                        <div className="flex items-center gap-2 flex-wrap">
                          <p className="text-sm font-bold text-slate-100">
                            {season.teamName || "Squadra"} <span className="text-slate-500">vs</span> {m.opponent}
                          </p>
                          {m.matchType && (
                            <Badge className={MATCH_TYPE_STYLES[m.matchType]}>
                              <TypeIcon className="w-3 h-3" /> {m.matchType}
                            </Badge>
                          )}
                          {m.coachNotes && (
                            <span title="Presenti annotazioni del mister">
                              <StickyNote className="w-3.5 h-3.5 text-amber-400" />
                            </span>
                          )}
                        </div>
                        <p className="text-xs text-slate-500">
                          {formatDate(m.date)} · {m.time} · {m.homeAway}
                          {m.matchType === "Torneo" && m.tournamentName ? ` · ${m.tournamentName}` : ""}
                        </p>
                      </div>
                    </div>
                    {m.status === "Disputata" && m.result ? (
                      <div className="text-lg font-extrabold text-slate-100">
                        {m.result.golFor} <span className="text-slate-600">-</span> {m.result.golAgainst}
                      </div>
                    ) : (
                      <Badge className="bg-sky-500/15 text-sky-400 border-sky-500/30">Da giocare</Badge>
                    )}
                  </div>
                </Card>
              </button>
            );
          })}
        </div>
      )}

      <Modal open={showAdd} onClose={() => setShowAdd(false)} title="Programma Partita" wide>
        <MatchForm onSubmit={addMatch} onCancel={() => setShowAdd(false)} players={players} teamColors={{ primary: season.colorPrimary, secondary: season.colorSecondary }} teamName={season.teamName} />
      </Modal>

      <Modal open={!!selected} onClose={() => setSelected(null)} title={selected ? `vs ${selected.opponent}` : ""} wide>
        {selected && (
          <MatchDetail
            match={matches.find((m) => m.id === selected.id) || selected}
            players={players}
            teamColors={{ primary: season.colorPrimary, secondary: season.colorSecondary }}
            teamName={season.teamName}
            lineup={season.lineup}
            clubLabel={season.teamName || ""}
            showToast={showToast}
            onUpdate={(patch) => {
              updateMatch(selected.id, patch);
              setSelected((prev) => ({ ...prev, ...patch }));
            }}
            onDelete={() => deleteMatch(selected.id)}
            onClose={() => setSelected(null)}
          />
        )}
      </Modal>
    </div>
  );
}

function MatchForm({ onSubmit, onCancel, players, teamColors, teamName, initial }) {
  const isEditing = !!initial;
  const [form, setForm] = useState(
    initial || {
      opponent: "",
      date: todayISO(),
      time: "15:00",
      homeAway: "Casa",
      venue: "",
      matchType: "Campionato",
      tournamentName: "",
      opponentColorPrimary: "#dc2626",
      opponentColorSecondary: "#0f172a",
      // Di default convocati tutti i giocatori disponibili (esclusi infortunati/squalificati)
      convocati: players.filter((p) => p.medicalStatus !== "Infortunato" && p.medicalStatus !== "Squalificato").map((p) => p.id),
    }
  );

  function toggleConvocato(id) {
    const list = form.convocati || [];
    setForm({ ...form, convocati: list.includes(id) ? list.filter((x) => x !== id) : [...list, id] });
  }

  return (
    <div>
      <Field label="Avversario">
        <input className={inputClass} value={form.opponent} onChange={(e) => setForm({ ...form, opponent: e.target.value })} placeholder="Nome squadra avversaria" />
      </Field>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Data">
          <input type="date" className={inputClass} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} />
        </Field>
        <Field label="Ora">
          <input type="time" className={inputClass} value={form.time} onChange={(e) => setForm({ ...form, time: e.target.value })} />
        </Field>
      </div>
      <Field label="Tipo partita">
        <div className="flex gap-2 flex-wrap">
          {MATCH_TYPES.map((t) => {
            const TypeIcon = MATCH_TYPE_ICONS[t];
            return (
              <button
                key={t}
                type="button"
                onClick={() => setForm({ ...form, matchType: t })}
                className={`flex-1 flex items-center justify-center gap-1.5 rounded-xl border px-3 py-2.5 text-sm font-medium ${
                  form.matchType === t ? "bg-emerald-500 text-slate-950 border-emerald-500" : "border-white/10 text-slate-400"
                }`}
              >
                <TypeIcon className="w-3.5 h-3.5" /> {t}
              </button>
            );
          })}
        </div>
      </Field>
      {form.matchType === "Torneo" && (
        <Field label="Nome torneo">
          <input
            className={inputClass}
            value={form.tournamentName}
            onChange={(e) => setForm({ ...form, tournamentName: e.target.value })}
            placeholder="Es. Torneo di Primavera"
          />
        </Field>
      )}
      <Field label="Casa / Trasferta">
        <div className="flex gap-2">
          {["Casa", "Trasferta"].map((o) => (
            <button
              key={o}
              type="button"
              onClick={() => setForm({ ...form, homeAway: o })}
              className={`flex-1 rounded-xl border px-3 py-2.5 text-sm font-medium ${
                form.homeAway === o ? "bg-emerald-500 text-slate-950 border-emerald-500" : "border-white/10 text-slate-400"
              }`}
            >
              {o}
            </button>
          ))}
        </div>
      </Field>
      <Field label="Luogo / Campo">
        <input className={inputClass} value={form.venue} onChange={(e) => setForm({ ...form, venue: e.target.value })} placeholder="Es. Campo Comunale" />
      </Field>

      <div className="grid sm:grid-cols-2 gap-x-4">
        <Field label="Colore primario avversario">
          <ColorSwatchPicker value={form.opponentColorPrimary} onChange={(hex) => setForm({ ...form, opponentColorPrimary: hex })} />
        </Field>
        <Field label="Colore secondario avversario">
          <ColorSwatchPicker value={form.opponentColorSecondary} onChange={(hex) => setForm({ ...form, opponentColorSecondary: hex })} />
        </Field>
      </div>

      <div className="flex items-center gap-2 mb-2 mt-1">
        <span className="text-[11px] text-slate-500">{teamName || "Squadra"}</span>
        <ColorPair primary={teamColors?.primary} secondary={teamColors?.secondary} />
        <span className="text-[11px] text-slate-500 ml-3">{form.opponent || "Avversario"}</span>
        <ColorPair primary={form.opponentColorPrimary} secondary={form.opponentColorSecondary} />
      </div>

      <div className="mt-3">
        <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide mb-2">
          Convocati ({(form.convocati || []).length})
        </p>
        {(!players || players.length === 0) ? (
          <p className="text-xs text-slate-500">Aggiungi giocatori alla rosa per poterli convocare.</p>
        ) : (
          <div className="overflow-y-auto pr-1" style={{ maxHeight: "42vh" }}>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, minmax(0, 1fr))", gap: "6px" }}>
              {players.map((p) => {
                const disabled = p.medicalStatus === "Infortunato" || p.medicalStatus === "Squalificato";
                const inDoubt = p.medicalStatus === "In dubbio";
                const checked = (form.convocati || []).includes(p.id);
                return (
                  <button
                    key={p.id}
                    type="button"
                    disabled={disabled}
                    onClick={() => toggleConvocato(p.id)}
                    className={`flex flex-col items-center gap-1 rounded-xl border p-2 text-center transition-colors ${
                      disabled
                        ? "border-rose-500/20 bg-rose-500/5 opacity-60 cursor-not-allowed"
                        : inDoubt
                        ? "border-rose-500/50 bg-rose-500/10"
                        : checked
                        ? "border-emerald-500/40 bg-emerald-500/10"
                        : "border-white/10"
                    }`}
                  >
                    <div className="relative">
                      <img src={playerAvatar(p)} className="w-9 h-9 rounded-full bg-slate-800 object-cover" />
                      {checked && !disabled && (
                        <span className="absolute -bottom-1 -right-1 bg-emerald-500 rounded-full p-0.5 border-2 border-slate-900">
                          <CheckCircle2 className="w-3 h-3 text-slate-950" />
                        </span>
                      )}
                    </div>
                    <p className={`text-[11px] leading-tight truncate w-full ${inDoubt ? "text-rose-300" : "text-slate-200"}`}>{p.name}</p>
                    <p className="text-[9px] text-slate-500 leading-tight truncate w-full">{p.role}</p>
                    {disabled && (
                      <Badge className={`${MEDICAL_COLORS[p.medicalStatus]} text-[9px] px-1.5 py-0`}>{p.medicalStatus}</Badge>
                    )}
                    {inDoubt && (
                      <Badge className="bg-rose-500/20 text-rose-300 border-rose-500/40 text-[9px] px-1.5 py-0">In dubbio</Badge>
                    )}
                  </button>
                );
              })}
            </div>
          </div>
        )}
      </div>

      <div className="flex justify-end gap-2 mt-4">
        <Button variant="secondary" onClick={onCancel}>Annulla</Button>
        <Button
          disabled={!form.opponent.trim() || (form.matchType === "Torneo" && !form.tournamentName.trim())}
          onClick={() => onSubmit(form)}
        >
          <Save className="w-4 h-4" /> {isEditing ? "Salva Modifiche" : "Programma"}
        </Button>
      </div>
    </div>
  );
}

function MatchDetail({ match, players, onUpdate, onDelete, onClose, teamColors, lineup, clubLabel, teamName, showToast }) {
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [editing, setEditing] = useState(false);
  const [scorerError, setScorerError] = useState("");
  const [assistError, setAssistError] = useState("");
  const isPlayed = match.status === "Disputata";

  const [golFor, setGolFor] = useState(match.result?.golFor ?? 0);
  const [golAgainst, setGolAgainst] = useState(match.result?.golAgainst ?? 0);

  function playerName(id) {
    return players.find((p) => p.id === id)?.name || "—";
  }

  function toggleConvocato(id) {
    const list = match.convocati || [];
    onUpdate({ convocati: list.includes(id) ? list.filter((x) => x !== id) : [...list, id] });
  }

  const convocatiPlayers = players.filter((p) => (match.convocati || []).includes(p.id));

  function addScorer() {
    const usedIds = (match.scorers || []).map((s) => s.playerId);
    const nextPlayer = convocatiPlayers.find((p) => !usedIds.includes(p.id));
    if (!nextPlayer) return;
    const currentTotal = (match.scorers || []).reduce((sum, s) => sum + (Number(s.goals) || 0), 0);
    if (currentTotal + 1 > golFor) {
      setScorerError(`Non puoi aggiungere altri marcatori: il totale supererebbe le ${golFor} reti segnate.`);
      return;
    }
    setScorerError("");
    onUpdate({ scorers: [...(match.scorers || []), { playerId: nextPlayer.id, goals: 1 }] });
  }
  function updateScorer(idx, patch) {
    const list = [...(match.scorers || [])];
    list[idx] = { ...list[idx], ...patch };
    if ("goals" in patch) {
      const total = list.reduce((sum, s) => sum + (Number(s.goals) || 0), 0);
      if (total > golFor) {
        setScorerError(`La somma dei marcatori (${total}) supera le ${golFor} reti segnate. Modifica non salvata.`);
        return;
      }
    }
    setScorerError("");
    onUpdate({ scorers: list });
  }
  function removeScorer(idx) {
    setScorerError("");
    onUpdate({ scorers: (match.scorers || []).filter((_, i) => i !== idx) });
  }

  function addAssist() {
    const usedIds = (match.assists || []).map((s) => s.playerId);
    const nextPlayer = convocatiPlayers.find((p) => !usedIds.includes(p.id));
    if (!nextPlayer) return;
    const currentTotal = (match.assists || []).reduce((sum, s) => sum + (Number(s.assists) || 0), 0);
    if (currentTotal + 1 > golFor) {
      setAssistError(`Non puoi aggiungere altri assist: il totale supererebbe le ${golFor} reti segnate.`);
      return;
    }
    setAssistError("");
    onUpdate({ assists: [...(match.assists || []), { playerId: nextPlayer.id, assists: 1 }] });
  }
  function updateAssist(idx, patch) {
    const list = [...(match.assists || [])];
    list[idx] = { ...list[idx], ...patch };
    if ("assists" in patch) {
      const total = list.reduce((sum, s) => sum + (Number(s.assists) || 0), 0);
      if (total > golFor) {
        setAssistError(`La somma degli assist (${total}) supera le ${golFor} reti segnate. Modifica non salvata.`);
        return;
      }
    }
    setAssistError("");
    onUpdate({ assists: list });
  }
  function removeAssist(idx) {
    setAssistError("");
    onUpdate({ assists: (match.assists || []).filter((_, i) => i !== idx) });
  }

  function toggleCard(type, id) {
    const key = type === "yellow" ? "yellowCards" : "redCards";
    const list = match[key] || [];
    onUpdate({ [key]: list.includes(id) ? list.filter((x) => x !== id) : [...list, id] });
  }

  const TypeIcon = MATCH_TYPE_ICONS[match.matchType] || Trophy;

  function handlePrintConvocation() {
    if (isPlayed) {
      // PARTITA DISPUTATA: risultato, marcatori/assist, ammoniti/espulsi, convocati
      let body = `<h1>${clubLabel || ""}</h1>`;
      body += `<h2>${match.homeAway === "Casa" ? clubLabel : match.opponent} ${match.result?.golFor ?? 0} — ${match.result?.golAgainst ?? 0} ${match.homeAway === "Casa" ? match.opponent : clubLabel}</h2>`;
      body += `<p>${match.matchType || ""} · ${formatDate(match.date)} · ${match.time} · ${match.homeAway}</p>`;
      body += `<p>${match.venue || ""}</p>`;
      if ((match.scorers || []).length > 0) {
        body += `<h2>Marcatori</h2><ul>${(match.scorers || [])
          .map((s) => `<li>${playerName(s.playerId)} (${s.goals})</li>`)
          .join("")}</ul>`;
      }
      if ((match.assists || []).length > 0) {
        body += `<h2>Assistman</h2><ul>${(match.assists || [])
          .map((s) => `<li>${playerName(s.playerId)} (${s.assists})</li>`)
          .join("")}</ul>`;
      }
      if ((match.yellowCards || []).length > 0 || (match.redCards || []).length > 0) {
        body += `<h2>Ammoniti / Espulsi</h2><ul>`;
        (match.yellowCards || []).forEach((id) => (body += `<li>${playerName(id)} — Ammonito</li>`));
        (match.redCards || []).forEach((id) => (body += `<li>${playerName(id)} — Espulso</li>`));
        body += `</ul>`;
      }
      body += `<h2>Convocati (${(match.convocati || []).length})</h2><ol>`;
      (match.convocati || []).forEach((id) => {
        const p = players.find((pl) => pl.id === id);
        if (p) body += `<li>#${p.number ?? "-"} ${p.name} — ${p.role}</li>`;
      });
      body += `</ol>`;
      if (match.coachNotes) {
        body += `<h2>Annotazioni mister</h2><p>${match.coachNotes}</p>`;
      }
      downloadPrintableHTML(`risultato-${match.opponent}-${match.date}.html`, "Risultato e Convocati", body);
    } else {
      // PARTITA PROGRAMMATA: solo dati partita e convocati, nessun modulo
      let body = `<h1>${clubLabel || ""}</h1>`;
      body += `<h2>Convocazione — vs ${match.opponent}</h2>`;
      body += `<p>${match.matchType || ""} · ${formatDate(match.date)} · ${match.time} · ${match.homeAway}</p>`;
      body += `<p>${match.venue || ""}</p>`;
      body += `<h2>Convocati (${(match.convocati || []).length})</h2><ol>`;
      (match.convocati || []).forEach((id) => {
        const p = players.find((pl) => pl.id === id);
        if (p) body += `<li>#${p.number ?? "-"} ${p.name} — ${p.role}</li>`;
      });
      body += `</ol>`;
      downloadPrintableHTML(`convocazione-${match.opponent}-${match.date}.html`, "Convocazione Partita", body);
    }
  }

  if (editing) {
    return (
      <MatchForm
        initial={{ ...match }}
        players={players}
        teamColors={teamColors}
        teamName={teamName}
        onCancel={() => setEditing(false)}
        onSubmit={(patch) => {
          onUpdate(patch);
          setEditing(false);
        }}
      />
    );
  }

  return (
    <div>
      <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
        <div>
          <div className="flex items-center gap-2 flex-wrap mb-1">
            {match.matchType && (
              <Badge className={MATCH_TYPE_STYLES[match.matchType]}>
                <TypeIcon className="w-3 h-3" /> {match.matchType}
              </Badge>
            )}
            <Badge className="bg-white/5 text-slate-300 border-white/10">{match.homeAway}</Badge>
          </div>
          <p className="text-sm text-slate-400">{formatDate(match.date)} · {match.time}</p>
          <p className="text-xs text-slate-500">{match.venue}</p>
          {match.matchType === "Torneo" && match.tournamentName && (
            <p className="text-xs text-amber-400 font-medium mt-0.5">🏆 {match.tournamentName}</p>
          )}
        </div>
        <div className="flex items-center gap-1.5 flex-wrap no-print">
          {!isPlayed && (
            <button
              onClick={async () => {
                const imageUrl = generateConvocationImage(match, players, clubLabel);
                const file = dataUrlToFile(imageUrl, `convocazione-${(match.opponent || "partita").replace(/[^a-z0-9]+/gi, "-")}.png`);
                const meetingTime = subtractMinutesFromTime(match.time, 45);
                const label = [
                  `Convocazione vs ${match.opponent || ""} - ${formatDate(match.date)} - ore ${match.time || "--"}${match.venue ? ` - campo ${match.venue}` : ""}`,
                  meetingTime
                    ? `Appuntamento al campo alle ore ${meetingTime}; presentarsi con tuta di rappresentanza, portare calzettoni e parastinchi`
                    : "Presentarsi con tuta di rappresentanza, portare calzettoni e parastinchi",
                ].join("\n");
                const shared = await shareFilesNative([file], { title: `Convocazione vs ${match.opponent || ""}`, text: label });
                if (!shared) {
                  const a = document.createElement("a");
                  a.href = imageUrl;
                  a.download = file.name;
                  document.body.appendChild(a);
                  a.click();
                  document.body.removeChild(a);
                  openWhatsAppFallback(`${label}\n(allega l'immagine appena scaricata)`);
                  showToast?.("Condivisione diretta non disponibile: immagine scaricata, allegala su WhatsApp manualmente.", "error");
                }
              }}
              className="rounded-lg p-2 hover:bg-white/10 text-slate-400"
              title="Condividi convocazione (es. su WhatsApp)"
            >
              <Share2 className="w-4 h-4" />
            </button>
          )}
          <button onClick={handlePrintConvocation} className="rounded-lg p-2 hover:bg-white/10 text-slate-400" title={isPlayed ? "Scarica risultato e convocati" : "Scarica modulo e convocazione (2 pagine)"}>
            <Download className="w-4 h-4" />
          </button>
          {!isPlayed && (
            <button onClick={() => setEditing(true)} className="rounded-lg p-2 hover:bg-white/10 text-slate-400" title="Modifica dati e convocati">
              <Edit2 className="w-4 h-4" />
            </button>
          )}
          <button onClick={() => setConfirmDelete(true)} className="rounded-lg p-2 hover:bg-rose-500/10 text-rose-400" title="Elimina">
            <Trash2 className="w-4 h-4" />
          </button>
          {onClose && (
            <Button onClick={onClose}>
              <ChevronLeft className="w-4 h-4" /> Salva e torna indietro
            </Button>
          )}
        </div>
      </div>

      <div className="flex items-center gap-2 mb-4 text-xs text-slate-500">
        <span>{teamName || "Squadra"}</span>
        <ColorPair primary={teamColors?.primary} secondary={teamColors?.secondary} size={16} />
        <span className="mx-1 text-slate-600">vs</span>
        <ColorPair primary={match.opponentColorPrimary} secondary={match.opponentColorSecondary} size={16} />
        <span>{match.opponent}</span>
      </div>

      {convocatiPlayers.length > 0 && (
        <div className="mb-5">
          <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide mb-2">Convocati ({convocatiPlayers.length})</p>
          <div className="rounded-xl border border-white/10 p-3 flex flex-wrap gap-x-4 gap-y-1.5">
            {convocatiPlayers.map((p) => (
              <span key={p.id} className="text-sm text-slate-300">
                <span className="font-medium text-slate-100">{p.name}</span> — {p.role}
              </span>
            ))}
          </div>
        </div>
      )}

      {confirmDelete && (
        <div className="rounded-xl border border-rose-500/30 bg-rose-500/10 p-3 mb-4 flex items-center justify-between gap-2">
          <p className="text-sm text-rose-300">Eliminare questa partita?</p>
          <div className="flex gap-2 shrink-0">
            <Button variant="secondary" onClick={() => setConfirmDelete(false)}>Annulla</Button>
            <Button variant="danger" onClick={onDelete}>Elimina</Button>
          </div>
        </div>
      )}

      {!isPlayed ? (
        <div className="mb-5">
          <Button
            onClick={() =>
              onUpdate({ status: "Disputata", result: { golFor: 0, golAgainst: 0 } })
            }
          >
            <CheckCircle2 className="w-4 h-4" /> Segna come Disputata
          </Button>
          <p className="text-xs text-slate-500 mt-2">Marca la partita come disputata per inserire il tabellino.</p>
        </div>
      ) : (
        <>
          <div className="flex items-center justify-center gap-4 mb-6">
            <input
              type="number"
              value={golFor}
              min="0"
              onChange={(e) => {
                const newGolFor = Number(e.target.value);
                const scorerTotal = (match.scorers || []).reduce((sum, s) => sum + (Number(s.goals) || 0), 0);
                const assistTotal = (match.assists || []).reduce((sum, s) => sum + (Number(s.assists) || 0), 0);
                if (newGolFor < scorerTotal || newGolFor < assistTotal) {
                  setScorerError(`Non puoi impostare ${newGolFor} reti: hai già assegnato ${scorerTotal} reti ai marcatori e ${assistTotal} agli assist. Riduci prima quelli.`);
                  return;
                }
                setScorerError("");
                setGolFor(newGolFor);
                onUpdate({ result: { ...match.result, golFor: newGolFor } });
              }}
              className="w-16 rounded-xl border border-white/10 bg-slate-950/60 text-center text-2xl font-extrabold py-2 text-emerald-400"
            />
            <span className="text-slate-600 text-xl">-</span>
            <input
              type="number"
              value={golAgainst}
              min="0"
              onChange={(e) => {
                setGolAgainst(Number(e.target.value));
                onUpdate({ result: { ...match.result, golAgainst: Number(e.target.value) } });
              }}
              className="w-16 rounded-xl border border-white/10 bg-slate-950/60 text-center text-2xl font-extrabold py-2 text-slate-100"
            />
          </div>

          {/* MARCATORI */}
          <div className="mb-5">
            <div className="flex items-center justify-between mb-2">
              <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide">Marcatori</p>
              <button onClick={addScorer} className="text-emerald-400 text-xs font-medium flex items-center gap-1">
                <Plus className="w-3.5 h-3.5" /> Aggiungi
              </button>
            </div>
            {(match.scorers || []).map((s, i) => {
              const usedByOthers = (match.scorers || []).filter((_, oi) => oi !== i).map((o) => o.playerId);
              const availableForRow = convocatiPlayers.filter((p) => p.id === s.playerId || !usedByOthers.includes(p.id));
              return (
                <div key={i} className="flex items-center gap-2 mb-2 flex-wrap">
                  <select className={inputClass} style={{ flex: "2 1 140px" }} value={s.playerId} onChange={(e) => updateScorer(i, { playerId: e.target.value })}>
                    {availableForRow.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
                  </select>
                  <input type="number" min="1" className={inputClass} style={{ width: 64 }} value={s.goals} onChange={(e) => updateScorer(i, { goals: Number(e.target.value) })} />
                  <button onClick={() => removeScorer(i)} className="text-rose-400 p-2"><X className="w-4 h-4" /></button>
                </div>
              );
            })}
            {scorerError && (
              <p className="text-[11px] text-rose-400 flex items-center gap-1 mt-1">
                <AlertCircle className="w-3.5 h-3.5 shrink-0" /> {scorerError}
              </p>
            )}
            {convocatiPlayers.length === 0 && (
              <p className="text-[11px] text-slate-500">Nessun convocato per questa partita: aggiungili modificando la partita.</p>
            )}
          </div>

          {/* ASSIST */}
          <div className="mb-5">
            <div className="flex items-center justify-between mb-2">
              <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide">Assistman</p>
              <button onClick={addAssist} className="text-emerald-400 text-xs font-medium flex items-center gap-1">
                <Plus className="w-3.5 h-3.5" /> Aggiungi
              </button>
            </div>
            {(match.assists || []).map((s, i) => {
              const usedByOthers = (match.assists || []).filter((_, oi) => oi !== i).map((o) => o.playerId);
              const availableForRow = convocatiPlayers.filter((p) => p.id === s.playerId || !usedByOthers.includes(p.id));
              return (
                <div key={i} className="flex items-center gap-2 mb-2">
                  <select className={inputClass} value={s.playerId} onChange={(e) => updateAssist(i, { playerId: e.target.value })}>
                    {availableForRow.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
                  </select>
                  <input type="number" min="1" className={inputClass + " w-20"} value={s.assists} onChange={(e) => updateAssist(i, { assists: Number(e.target.value) })} />
                  <button onClick={() => removeAssist(i)} className="text-rose-400 p-2"><X className="w-4 h-4" /></button>
                </div>
              );
            })}
            {assistError && (
              <p className="text-[11px] text-rose-400 flex items-center gap-1 mt-1">
                <AlertCircle className="w-3.5 h-3.5 shrink-0" /> {assistError}
              </p>
            )}
          </div>

          {/* AMMONITI / ESPULSI + CONVOCATI */}
          <div className="mb-2">
            <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide mb-2">Convocati, ammoniti ed espulsi</p>
            <div className="space-y-2 overflow-y-auto pr-1" style={{ maxHeight: "40vh" }}>
              {players.map((p) => (
                <div key={p.id} className="flex items-center justify-between gap-2 rounded-xl border border-white/10 p-2.5">
                  <div className="flex items-center gap-2 min-w-0">
                    <img src={playerAvatar(p)} className="w-8 h-8 rounded-full bg-slate-800 object-cover" />
                    <p className="text-sm text-slate-200 truncate">{p.name}</p>
                  </div>
                  <div className="flex items-center gap-1.5 shrink-0">
                    <button
                      onClick={() => toggleConvocato(p.id)}
                      className={`rounded-full border px-2 py-1 text-[10px] font-medium ${
                        (match.convocati || []).includes(p.id) ? "bg-sky-500/15 text-sky-400 border-sky-500/30" : "border-white/10 text-slate-600"
                      }`}
                    >
                      Convocato
                    </button>
                    <button
                      onClick={() => toggleCard("yellow", p.id)}
                      className={`w-5 h-6 rounded-sm border ${
                        (match.yellowCards || []).includes(p.id) ? "bg-amber-400 border-amber-400" : "border-white/20"
                      }`}
                      title="Ammonizione"
                    />
                    <button
                      onClick={() => toggleCard("red", p.id)}
                      className={`w-5 h-6 rounded-sm border ${
                        (match.redCards || []).includes(p.id) ? "bg-rose-500 border-rose-500" : "border-white/20"
                      }`}
                      title="Espulsione"
                    />
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* ANNOTAZIONI MISTER */}
          <div className="mt-5">
            <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide mb-2 flex items-center gap-1.5">
              <StickyNote className="w-3.5 h-3.5" /> Annotazioni mister
            </p>
            <textarea
              rows={3}
              className={inputClass}
              value={match.coachNotes || ""}
              onChange={(e) => onUpdate({ coachNotes: e.target.value })}
              placeholder="Note su formazione, episodi, prestazioni individuali..."
            />
          </div>
        </>
      )}
    </div>
  );
}

/* ============================================================
   SEZIONE MODULI IN CAMPO
   ============================================================ */

function StrengthsWeaknesses({ formation, compact }) {
  return (
    <div className={`grid sm:grid-cols-2 gap-4 ${compact ? "" : "mt-2"}`}>
      <div className="rounded-xl border border-emerald-500/20 bg-emerald-500/5 p-3.5">
        <p className="text-xs font-bold text-emerald-400 uppercase tracking-wide mb-2">Punti di forza</p>
        <ul className="space-y-1.5">
          {formation.strengths.map((s, i) => (
            <li key={i} className="text-xs text-slate-300 flex gap-1.5">
              <span className="text-emerald-400 shrink-0">+</span>
              <span>{s}</span>
            </li>
          ))}
        </ul>
      </div>
      <div className="rounded-xl border border-rose-500/20 bg-rose-500/5 p-3.5">
        <p className="text-xs font-bold text-rose-400 uppercase tracking-wide mb-2">Punti di debolezza</p>
        <ul className="space-y-1.5">
          {formation.weaknesses.map((s, i) => (
            <li key={i} className="text-xs text-slate-300 flex gap-1.5">
              <span className="text-rose-400 shrink-0">−</span>
              <span>{s}</span>
            </li>
          ))}
        </ul>
      </div>
    </div>
  );
}

// Legenda delle sigle di posizione usate nei moduli (es. DC = Difensore Centrale).
// Mostra solo le sigle effettivamente presenti nel modulo corrente, se fornito;
// altrimenti l'elenco completo.
function PositionLegend({ formation }) {
  const labelsInUse = formation ? new Set((formation.positions || []).map((p) => p.label)) : null;
  const entries = labelsInUse
    ? FORMATION_POSITION_LIBRARY.filter((p) => labelsInUse.has(p.label))
    : FORMATION_POSITION_LIBRARY;
  return (
    <div className="rounded-xl border border-white/10 bg-white/5 p-3.5 mt-4">
      <p className="text-xs font-bold text-slate-300 uppercase tracking-wide mb-2.5">Legenda posizioni</p>
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-3 gap-y-1.5">
        {entries.map((p) => (
          <p key={p.label} className="text-[11px] text-slate-400">
            <span className="font-mono font-bold text-slate-200">{p.label}</span> = {p.description}
          </p>
        ))}
      </div>
    </div>
  );
}

function FormationPickerCard({ formation, onChoose }) {
  return (
    <Card className="p-4 flex flex-col">
      <div className="mb-2">
        <p className="text-lg font-extrabold text-slate-100">{formation.name}</p>
        <p className="text-xs font-semibold text-emerald-400 uppercase tracking-wide">{formation.subtitle}</p>
      </div>
      <p className="text-xs text-slate-400 mb-2">{formation.description}</p>
      <p className="text-[11px] font-mono text-slate-500 bg-slate-950/60 rounded-lg px-2.5 py-2 mb-3 break-words">
        {formation.structureLabel}
      </p>
      <StrengthsWeaknesses formation={formation} compact />
      {formation.note && (
        <p className="text-[11px] text-amber-300/90 bg-amber-500/10 border border-amber-500/20 rounded-lg px-2.5 py-2 mt-3 italic">
          {formation.note}
        </p>
      )}
      <Button className="mt-4" onClick={() => onChoose(formation.id)}>
        <CheckCircle2 className="w-4 h-4" /> Scegli questo modulo
      </Button>
    </Card>
  );
}

function PitchBackground() {
  return (
    <svg viewBox="0 0 68 100" preserveAspectRatio="none" className="absolute inset-0 w-full h-full">
      <defs>
        <linearGradient id="grassGradient" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor="#14532d" />
          <stop offset="100%" stopColor="#166534" />
        </linearGradient>
      </defs>
      <rect x="0" y="0" width="68" height="100" fill="url(#grassGradient)" />
      {[0, 1, 2, 3, 4, 5, 6, 7].map((i) => (
        <rect key={i} x="0" y={i * 12.5} width="68" height="12.5" fill={i % 2 === 0 ? "#ffffff08" : "transparent"} />
      ))}
      <g stroke="#ffffff70" strokeWidth="0.4" fill="none">
        <rect x="2" y="2" width="64" height="96" />
        <line x1="2" y1="50" x2="66" y2="50" />
        <circle cx="34" cy="50" r="9" />
        <circle cx="34" cy="50" r="0.6" fill="#ffffff70" />
        {/* area di rigore bassa (portiere) */}
        <rect x="14" y="80" width="40" height="18" />
        <rect x="24" y="90" width="20" height="8" />
        <circle cx="34" cy="86" r="0.6" fill="#ffffff70" />
        {/* area di rigore alta */}
        <rect x="14" y="2" width="40" height="18" />
        <rect x="24" y="2" width="20" height="8" />
        <circle cx="34" cy="14" r="0.6" fill="#ffffff70" />
      </g>
    </svg>
  );
}

function Pitch({ formation, assignments, players, onSlotClick }) {
  return (
    <div className="relative w-full rounded-2xl overflow-hidden border border-white/10 shadow-xl" style={{ aspectRatio: "68 / 100" }}>
      <PitchBackground />
      {formation.positions.map((pos) => {
        const player = players.find((p) => p.id === assignments[pos.id]);
        return (
          <button
            key={pos.id}
            onClick={() => onSlotClick(pos)}
            style={{ left: `${pos.x}%`, top: `${pos.y}%` }}
            className="absolute -translate-x-1/2 -translate-y-1/2 flex flex-col items-center"
          >
            {player ? (
              <>
                <img
                  src={playerAvatar(player)}
                  alt={player.name}
                  className="w-10 h-10 sm:w-12 sm:h-12 rounded-full border-2 border-emerald-400 object-cover shadow-lg bg-slate-800"
                />
                <span className="mt-1 text-[9px] sm:text-[10px] font-bold text-white bg-slate-950/85 px-1.5 py-0.5 rounded whitespace-nowrap max-w-[76px] truncate">
                  {player.number != null ? `${player.number} · ` : ""}{shortName(player.name)}
                </span>
              </>
            ) : (
              <>
                <div className="w-10 h-10 sm:w-12 sm:h-12 rounded-full border-2 border-dashed border-white/50 bg-slate-950/40 flex items-center justify-center text-white text-[9px] sm:text-[10px] font-bold">
                  {pos.label}
                </div>
                <span className="mt-1 text-[9px] sm:text-[10px] text-white/70">Scegli</span>
              </>
            )}
          </button>
        );
      })}
    </div>
  );
}

function PlayerPickerModal({ open, onClose, title, players, excludedIds, currentId, suggestedRole, onSelect, onClear, onUpdateStatus }) {
  const config = useConfig();
  const [editingStatusId, setEditingStatusId] = useState(null);
  const [pendingStatus, setPendingStatus] = useState("Disponibile");

  function startEditStatus(p) {
    setEditingStatusId(p.id);
    setPendingStatus(p.medicalStatus);
  }

  function confirmStatus(playerId) {
    onUpdateStatus(playerId, pendingStatus);
    setEditingStatusId(null);
  }

  return (
    <Modal open={open} onClose={onClose} title={title}>
      {currentId && onClear && (
        <button
          onClick={onClear}
          className="w-full mb-3 flex items-center justify-center gap-2 rounded-xl border border-rose-500/30 bg-rose-500/10 p-2.5 text-sm font-medium text-rose-400 hover:bg-rose-500/15"
        >
          <X className="w-4 h-4" /> Rimuovi assegnazione
        </button>
      )}
      {players.length === 0 ? (
        <EmptyState icon={Users} text="Nessun giocatore in rosa." />
      ) : (
        <div className="space-y-2 overflow-y-auto pr-1" style={{ maxHeight: "55vh" }}>
          {[...players]
            .filter((p) => p.id === currentId || !excludedIds.has(p.id))
            .sort((a, b) => {
              const aMatch = suggestedRole && (a.role === suggestedRole || a.role2 === suggestedRole) ? 0 : 1;
              const bMatch = suggestedRole && (b.role === suggestedRole || b.role2 === suggestedRole) ? 0 : 1;
              if (aMatch !== bMatch) return aMatch - bMatch;
              return a.name.localeCompare(b.name);
            })
            .map((p) => {
              const disabled = p.medicalStatus === "Infortunato" || p.medicalStatus === "Squalificato";
              const inDoubt = p.medicalStatus === "In dubbio";
              const isEditingStatus = editingStatusId === p.id;
              return (
                <div
                  key={p.id}
                  className={`w-full rounded-xl border p-2.5 transition-colors ${
                    disabled
                      ? "border-rose-500/20 bg-rose-500/5"
                      : inDoubt
                      ? "border-rose-500/50 bg-rose-500/10"
                      : p.id === currentId
                      ? "border-emerald-500/40 bg-white/5"
                      : "border-white/10"
                  }`}
                >
                  <div className="flex items-center gap-3">
                    <button
                      disabled={disabled}
                      onClick={() => onSelect(p.id)}
                      className={`flex-1 flex items-center gap-3 text-left min-w-0 ${disabled ? "opacity-60 cursor-not-allowed" : ""}`}
                    >
                      <img src={playerAvatar(p)} alt={p.name} className="w-9 h-9 rounded-full bg-slate-800 object-cover shrink-0" />
                      <div className="flex-1 min-w-0">
                        <p className={`text-sm font-medium truncate ${inDoubt ? "text-rose-300" : "text-slate-200"}`}>{p.name}</p>
                        <div className="flex items-center gap-1.5 mt-0.5 flex-wrap">
                          <Badge className={ROLE_COLORS[p.role]}>{p.role}</Badge>
                          {p.role !== "Portiere" && p.position && (
                            <Badge className={ROLE_COLORS[p.role]}>{p.position}</Badge>
                          )}
                          {p.role2 && <Badge className="bg-white/10 text-slate-200 border-white/20">Alt: {p.role2}</Badge>}
                          {disabled && (
                            <Badge className={MEDICAL_COLORS[p.medicalStatus]}>
                              <HeartPulse className="w-3 h-3" /> {p.medicalStatus}
                            </Badge>
                          )}
                          {inDoubt && (
                            <Badge className="bg-rose-500/20 text-rose-300 border-rose-500/40">
                              <AlertCircle className="w-3 h-3" /> In dubbio
                            </Badge>
                          )}
                        </div>
                      </div>
                    </button>
                    {onUpdateStatus && !isEditingStatus && (
                      <button
                        onClick={() => startEditStatus(p)}
                        className="rounded-lg p-1.5 hover:bg-white/10 text-slate-400 shrink-0"
                        title="Modifica stato medico"
                      >
                        <Edit2 className="w-3.5 h-3.5" />
                      </button>
                    )}
                  </div>
                  {isEditingStatus && (
                    <div className="mt-2.5 pt-2.5 border-t border-white/10">
                      <p className="text-[11px] text-slate-500 mb-1.5">Cambia stato medico di {p.name}:</p>
                      <div className="flex items-center gap-2 flex-wrap">
                        <select
                          className={inputClass + " flex-1"}
                          style={{ minWidth: 140 }}
                          value={pendingStatus}
                          onChange={(e) => setPendingStatus(e.target.value)}
                        >
                          {config.medicalStatuses.map((m) => (
                            <option key={m} value={m}>{m}</option>
                          ))}
                        </select>
                        <Button variant="secondary" className="px-3 py-2 text-xs" onClick={() => setEditingStatusId(null)}>
                          Annulla
                        </Button>
                        <Button className="px-3 py-2 text-xs" onClick={() => confirmStatus(p.id)}>
                          Conferma
                        </Button>
                      </div>
                    </div>
                  )}
                </div>
              );
            })}
        </div>
      )}
    </Modal>
  );
}

function emptyNewFormation() {
  return {
    format: "9v9",
    name: "",
    subtitle: "",
    positionLabels: [],
    strengths: [""],
    weaknesses: [""],
    note: "",
  };
}

function roleForPositionLabel(label) {
  return FORMATION_POSITION_LIBRARY.find((p) => p.label === label)?.role || "Centrocampista";
}

// Campo cliccabile per posizionare in sequenza le posizioni del nuovo modulo.
const PLACEMENT_GRID_COLS = 5; // colonne per il posizionamento, dalla linea laterale sinistra a quella destra
const PLACEMENT_GRID_ROWS = 8; // 4 righe per ciascuna metà campo
// Il rettangolo di gioco in PitchBackground occupa x: 2–66 (su viewBox largo 68) e y: 2–98 (su viewBox alto 100):
// la griglia deve agganciarsi esattamente a queste linee laterali, non ai bordi esterni dell'SVG.
const PITCH_X_MIN = (2 / 68) * 100;
const PITCH_X_MAX = (66 / 68) * 100;
const PITCH_Y_MIN = 2;
const PITCH_Y_MAX = 98;

function ClickToPlacePitch({ placed, onPlace, disabled }) {
  const containerRef = React.useRef(null);

  function handleClick(e) {
    if (disabled || !containerRef.current) return;
    const rect = containerRef.current.getBoundingClientRect();
    const relX = ((e.clientX - rect.left) / rect.width) * 100;
    const relY = ((e.clientY - rect.top) / rect.height) * 100;
    // Aggancia il click al centro del quadrante della griglia più vicino,
    // calcolato dentro i confini reali del rettangolo di gioco.
    const pitchWidth = PITCH_X_MAX - PITCH_X_MIN;
    const pitchHeight = PITCH_Y_MAX - PITCH_Y_MIN;
    const col = Math.min(PLACEMENT_GRID_COLS - 1, Math.max(0, Math.floor(((relX - PITCH_X_MIN) / pitchWidth) * PLACEMENT_GRID_COLS)));
    const row = Math.min(PLACEMENT_GRID_ROWS - 1, Math.max(0, Math.floor(((relY - PITCH_Y_MIN) / pitchHeight) * PLACEMENT_GRID_ROWS)));
    const x = PITCH_X_MIN + ((col + 0.5) / PLACEMENT_GRID_COLS) * pitchWidth;
    const y = PITCH_Y_MIN + ((row + 0.5) / PLACEMENT_GRID_ROWS) * pitchHeight;
    onPlace(x, y);
  }

  return (
    <div
      ref={containerRef}
      onClick={handleClick}
      className={`relative w-full rounded-2xl overflow-hidden border border-white/10 shadow-xl ${disabled ? "" : "cursor-crosshair"}`}
      style={{ aspectRatio: "68 / 100" }}
    >
      <PitchBackground />
      {/* Griglia guida: 8 righe x 5 colonne, agganciata alle linee laterali del campo */}
      <svg viewBox="0 0 100 100" preserveAspectRatio="none" className="absolute inset-0 w-full h-full pointer-events-none">
        <g stroke="#ffffff30" strokeWidth="0.3">
          {Array.from({ length: PLACEMENT_GRID_COLS - 1 }).map((_, i) => {
            const gx = PITCH_X_MIN + ((i + 1) / PLACEMENT_GRID_COLS) * (PITCH_X_MAX - PITCH_X_MIN);
            return <line key={`c${i}`} x1={gx} y1={PITCH_Y_MIN} x2={gx} y2={PITCH_Y_MAX} />;
          })}
          {Array.from({ length: PLACEMENT_GRID_ROWS - 1 }).map((_, i) => {
            const gy = PITCH_Y_MIN + ((i + 1) / PLACEMENT_GRID_ROWS) * (PITCH_Y_MAX - PITCH_Y_MIN);
            return <line key={`r${i}`} x1={PITCH_X_MIN} y1={gy} x2={PITCH_X_MAX} y2={gy} />;
          })}
        </g>
      </svg>
      {placed.map((p, i) => (
        <div
          key={p.id}
          style={{ left: `${p.x}%`, top: `${p.y}%` }}
          className="absolute -translate-x-1/2 -translate-y-1/2 flex flex-col items-center pointer-events-none"
        >
          <div className="w-9 h-9 rounded-full border-2 border-emerald-400 bg-slate-950/85 flex items-center justify-center text-[10px] font-bold text-white">
            {p.label}
          </div>
          <span className="mt-0.5 text-[9px] text-white/70">{i + 1}</span>
        </div>
      ))}
    </div>
  );
}

function NewFormationWizard({ onCancel, onSave }) {
  const [step, setStep] = useState(1);
  const [form, setForm] = useState(emptyNewFormation());
  const [placed, setPlaced] = useState([]);

  const requiredCount = FORMAT_PLAYER_COUNT[form.format];

  useEffect(() => {
    setForm((f) => {
      const count = FORMAT_PLAYER_COUNT[f.format];
      const labels = [...f.positionLabels];
      while (labels.length < count) labels.push("");
      labels.length = count;
      return { ...f, positionLabels: labels };
    });
  }, [form.format]);

  function updatePositionLabel(idx, label) {
    const next = [...form.positionLabels];
    next[idx] = label;
    setForm({ ...form, positionLabels: next });
  }

  function updateListItem(key, idx, value) {
    const next = [...form[key]];
    next[idx] = value;
    setForm({ ...form, [key]: next });
  }

  function addListItem(key) {
    if (form[key].length >= 6) return;
    setForm({ ...form, [key]: [...form[key], ""] });
  }

  function removeListItem(key, idx) {
    setForm({ ...form, [key]: form[key].filter((_, i) => i !== idx) });
  }

  const step1Valid = form.name.trim() && form.positionLabels.length === requiredCount && form.positionLabels.every((l) => l);

  function goToPlacement() {
    setPlaced([]);
    setStep(2);
  }

  function placeNext(x, y) {
    if (placed.length >= requiredCount) return;
    const label = form.positionLabels[placed.length];
    setPlaced([...placed, { id: uid("pos"), label, role: roleForPositionLabel(label), x, y }]);
  }

  function undoLastPlacement() {
    setPlaced(placed.slice(0, -1));
  }

  function handleSave() {
    const structureLabel = form.positionLabels.join(" — ");
    onSave({
      id: uid("customf"),
      format: form.format,
      name: form.name.trim(),
      subtitle: form.subtitle.trim(),
      description: form.subtitle.trim(),
      structureLabel,
      strengths: form.strengths.map((s) => s.trim()).filter(Boolean),
      weaknesses: form.weaknesses.map((s) => s.trim()).filter(Boolean),
      note: form.note.trim() || undefined,
      positions: placed,
      custom: true,
    });
  }

  if (step === 2) {
    const done = placed.length >= requiredCount;
    return (
      <div>
        <p className="text-sm text-slate-400 mb-4">
          {done
            ? "Tutte le posizioni sono state piazzate. Controlla e salva il modulo."
            : `Clicca sul campo per posizionare: ${form.positionLabels[placed.length]} (${placed.length + 1} di ${requiredCount})`}
        </p>
        <div style={{ maxWidth: 320 }} className="mx-auto">
          <ClickToPlacePitch placed={placed} onPlace={placeNext} disabled={done} />
        </div>
        <div className="flex justify-between gap-2 mt-5">
          <Button variant="secondary" onClick={() => setStep(1)}>
            <ChevronLeft className="w-4 h-4" /> Torna al modulo
          </Button>
          <div className="flex gap-2">
            {placed.length > 0 && (
              <Button variant="secondary" onClick={undoLastPlacement}>Annulla ultimo</Button>
            )}
            <Button disabled={!done} onClick={handleSave}>
              <Save className="w-4 h-4" /> Salva modulo
            </Button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div>
      <Field label="Numero giocatori in campo">
        <select className={inputClass} value={form.format} onChange={(e) => setForm({ ...form, format: e.target.value })}>
          {TEAM_FORMAT_OPTIONS.map((fmt) => (
            <option key={fmt} value={fmt}>Calcio a {fmt}</option>
          ))}
        </select>
      </Field>
      <Field label="Schema del modulo">
        <input className={inputClass} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="Es. 4-3-3" />
      </Field>
      <Field label="Descrizione modulo">
        <input className={inputClass} value={form.subtitle} onChange={(e) => setForm({ ...form, subtitle: e.target.value })} placeholder="Es. Il modulo del coraggio" />
      </Field>

      <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide mb-2 mt-4">
        Posizioni ({requiredCount} giocatori, portiere incluso)
      </p>
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-2 mb-4">
        {form.positionLabels.map((label, idx) => (
          <select
            key={idx}
            className={inputClass}
            value={label}
            onChange={(e) => updatePositionLabel(idx, e.target.value)}
          >
            <option value="">Posizione {idx + 1}...</option>
            {FORMATION_POSITION_LIBRARY.map((p) => (
              <option key={p.label} value={p.label}>{p.label} · {p.description}</option>
            ))}
          </select>
        ))}
      </div>

      <div className="grid sm:grid-cols-2 gap-4 mb-4">
        <div>
          <p className="text-xs font-bold text-emerald-400 uppercase tracking-wide mb-2">Punti di forza (max 6)</p>
          {form.strengths.map((s, i) => (
            <div key={i} className="flex gap-1.5 mb-1.5">
              <input className={inputClass} value={s} onChange={(e) => updateListItem("strengths", i, e.target.value)} placeholder="Punto di forza" />
              <button onClick={() => removeListItem("strengths", i)} className="text-rose-400 px-1"><X className="w-4 h-4" /></button>
            </div>
          ))}
          {form.strengths.length < 6 && (
            <button onClick={() => addListItem("strengths")} className="text-xs text-emerald-400 flex items-center gap-1 mt-1">
              <Plus className="w-3.5 h-3.5" /> Aggiungi
            </button>
          )}
        </div>
        <div>
          <p className="text-xs font-bold text-rose-400 uppercase tracking-wide mb-2">Punti di debolezza (max 6)</p>
          {form.weaknesses.map((s, i) => (
            <div key={i} className="flex gap-1.5 mb-1.5">
              <input className={inputClass} value={s} onChange={(e) => updateListItem("weaknesses", i, e.target.value)} placeholder="Punto di debolezza" />
              <button onClick={() => removeListItem("weaknesses", i)} className="text-rose-400 px-1"><X className="w-4 h-4" /></button>
            </div>
          ))}
          {form.weaknesses.length < 6 && (
            <button onClick={() => addListItem("weaknesses")} className="text-xs text-rose-400 flex items-center gap-1 mt-1">
              <Plus className="w-3.5 h-3.5" /> Aggiungi
            </button>
          )}
        </div>
      </div>

      <Field label="Nota aggiuntiva (facoltativa)">
        <textarea
          rows={2}
          className={inputClass}
          value={form.note}
          onChange={(e) => setForm({ ...form, note: e.target.value })}
          placeholder="Es. Consigliato come punto di partenza per la categoria Esordienti..."
        />
      </Field>

      <div className="flex justify-end gap-2 mt-4">
        <Button variant="secondary" onClick={onCancel}>Annulla</Button>
        <Button disabled={!step1Valid} onClick={goToPlacement}>
          Continua: posiziona sul campo <ChevronRight className="w-4 h-4" />
        </Button>
      </div>
    </div>
  );
}

function FormationsSection({ season, updateSeason, library, updateLibrary, showToast }) {
  const lineup = season.lineup || emptyLineup();
  const players = season.players || [];
  const allFormations = useMemo(() => [...FORMATIONS, ...(library.customFormations || [])], [library.customFormations]);
  const allFormationsByFormat = useMemo(() => {
    const custom = library.customFormations || [];
    const result = {};
    Object.keys(FORMATIONS_BY_FORMAT).forEach((fmt) => {
      result[fmt] = [...FORMATIONS_BY_FORMAT[fmt], ...custom.filter((f) => f.format === fmt)];
    });
    return result;
  }, [library.customFormations]);
  const formation = allFormations.find((f) => f.id === lineup.formationId);
  const [picker, setPicker] = useState(null); // { kind: 'position'|'bench', positionId?, benchIndex?, suggestedRole? }
  const [showNewFormation, setShowNewFormation] = useState(false);
  const [selectedFormat, setSelectedFormat] = useState(() => {
    if (lineup.formationId) {
      const found = Object.entries(allFormationsByFormat).find(([, arr]) => arr.some((f) => f.id === lineup.formationId));
      if (found) return found[0];
    }
    return TEAM_FORMAT_OPTIONS.includes(season.teamFormat) ? season.teamFormat : "9v9";
  });

  function saveCustomFormation(newFormation) {
    updateLibrary((lib) => ({ customFormations: [...(lib.customFormations || []), newFormation] }));
    setShowNewFormation(false);
    showToast("Modulo personalizzato creato");
  }

  function chooseFormation(id) {
    updateSeason(() => ({ lineup: { formationId: id, assignments: {}, bench: [] } }));
  }

  function changeFormation() {
    updateSeason(() => ({ lineup: emptyLineup() }));
    showToast("Modulo azzerato: scegline uno nuovo");
  }

  const assignedIds = new Set([
    ...Object.values(lineup.assignments || {}).filter(Boolean),
    ...(lineup.bench || []).filter(Boolean),
  ]);

  function assignToPosition(positionId, playerId) {
    updateSeason((s) => ({
      lineup: { ...(s.lineup || emptyLineup()), assignments: { ...((s.lineup || emptyLineup()).assignments), [positionId]: playerId } },
    }));
    setPicker(null);
  }

  function clearPosition(positionId) {
    updateSeason((s) => {
      const a = { ...((s.lineup || emptyLineup()).assignments) };
      delete a[positionId];
      return { lineup: { ...(s.lineup || emptyLineup()), assignments: a } };
    });
  }

  function updatePlayerStatus(playerId, status) {
    updateSeason((s) => ({
      players: (s.players || []).map((p) => (p.id === playerId ? { ...p, medicalStatus: status } : p)),
    }));
    showToast("Stato medico aggiornato");
  }

  function addToBench(playerId) {
    updateSeason((s) => {
      const bench = [...((s.lineup || emptyLineup()).bench || [])];
      if (bench.length >= 6) return {};
      bench.push(playerId);
      return { lineup: { ...(s.lineup || emptyLineup()), bench } };
    });
    setPicker(null);
  }

  function removeFromBench(index) {
    updateSeason((s) => {
      const bench = [...((s.lineup || emptyLineup()).bench || [])];
      bench.splice(index, 1);
      return { lineup: { ...(s.lineup || emptyLineup()), bench } };
    });
  }

  function handlePrint() {
    if (!formation) return;
    let body = `<h1>${season.teamName || ""}</h1>`;
    body += `<h2>${formation.name} — ${formation.subtitle}</h2>`;
    body += `<p>${formation.structureLabel}</p>`;
    body += buildPitchHTML(formation, lineup.assignments || {}, players);
    if ((lineup.bench || []).length > 0) {
      body += `<h2>Panchina</h2><ul>${(lineup.bench || [])
        .map((id) => {
          const p = players.find((pl) => pl.id === id);
          return p ? `<li>${p.name} — ${p.role}</li>` : "";
        })
        .join("")}</ul>`;
    }
    body += `<h2>Punti di forza</h2><ul>${formation.strengths.map((s) => `<li>${s}</li>`).join("")}</ul>`;
    body += `<h2>Punti di debolezza</h2><ul>${formation.weaknesses.map((s) => `<li>${s}</li>`).join("")}</ul>`;
    downloadPrintableHTML(`modulo-${formation.name}.html`, "Modulo in campo", body);
  }

  return (
    <div>
      <SectionTitle
        eyebrow="Tattica"
        title="Moduli in campo"
        icon={LayoutGrid}
        action={
          formation && (
            <div className="flex gap-2 no-print">
              <Button variant="secondary" onClick={changeFormation}>
                <ArrowLeftRight className="w-4 h-4" /> Cambia modulo
              </Button>
              <Button onClick={handlePrint}>
                <Download className="w-4 h-4" /> Stampa / Esporta PDF
              </Button>
            </div>
          )
        }
      />

      {!formation && (
        <div className="flex items-center justify-between gap-3 mb-5 flex-wrap">
          <div className="flex gap-2 flex-wrap">
            {TEAM_FORMAT_OPTIONS.map((fmt) => (
              <button
                key={fmt}
                onClick={() => setSelectedFormat(fmt)}
                className={`rounded-full px-4 py-2 text-xs font-semibold border transition-colors ${
                  selectedFormat === fmt ? "bg-emerald-500 text-slate-950 border-emerald-500" : "border-white/10 text-slate-400"
                }`}
              >
                Calcio a {fmt}
              </button>
            ))}
          </div>
          <Button variant="secondary" onClick={() => setShowNewFormation(true)}>
            <Plus className="w-4 h-4" /> Crea nuovo modulo
          </Button>
        </div>
      )}

      {!formation ? (
        <div>
          <p className="text-sm text-slate-400 mb-5">
            Scegli il modulo tattico per il campo a {selectedFormat} (portiere + giocatori di movimento). Ogni modulo mostra punti di forza e di debolezza per aiutarti nella scelta.
          </p>
          <div className="grid lg:grid-cols-2 gap-4">
            {allFormationsByFormat[selectedFormat].map((f) => (
              <FormationPickerCard key={f.id} formation={f} onChoose={chooseFormation} />
            ))}
          </div>
        </div>
      ) : (
        <div id="print-lineup">
          <div className="mb-4">
            <p className="text-xl font-extrabold text-slate-100">
              {formation.name} <span className="text-slate-500 font-normal">— {formation.subtitle}</span>
            </p>
            <p className="text-xs font-mono text-slate-500 mt-1">{formation.structureLabel}</p>
          </div>

          <div style={{ display: "flex", flexWrap: "wrap", gap: "24px" }}>
            <div style={{ flex: "0 0 320px", maxWidth: "380px" }}>
              <Pitch
                formation={formation}
                assignments={lineup.assignments || {}}
                players={players}
                onSlotClick={(pos) =>
                  setPicker({ kind: "position", positionId: pos.id, suggestedRole: pos.role, currentId: (lineup.assignments || {})[pos.id] })
                }
              />
              <p className="text-[11px] text-slate-500 mt-2 no-print">Tocca un giocatore sul campo per assegnare o sostituire.</p>
            </div>

            <div style={{ flex: "1 1 320px", minWidth: "0" }}>
              <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide mb-2">Panchina (max 6 riserve)</p>
              <div className="grid grid-cols-2 sm:grid-cols-3 gap-2 mb-6">
                {(lineup.bench || []).map((playerId, idx) => {
                  const p = players.find((pl) => pl.id === playerId);
                  return (
                    <div key={idx} className="rounded-xl border border-white/10 p-2.5 flex items-center gap-2 relative">
                      <img src={playerAvatar(p)} alt={p?.name} className="w-8 h-8 rounded-full bg-slate-800 object-cover shrink-0" />
                      <div className="min-w-0 flex-1">
                        <p className="text-xs font-medium text-slate-200 truncate">{p?.name || "—"}</p>
                        {p && <Badge className={`${ROLE_COLORS[p.role]} mt-0.5`}>{p.role}</Badge>}
                      </div>
                      <button
                        onClick={() => removeFromBench(idx)}
                        className="absolute top-1 right-1 text-rose-400 p-0.5 no-print"
                      >
                        <X className="w-3 h-3" />
                      </button>
                    </div>
                  );
                })}
                {(lineup.bench || []).length < 6 && (
                  <button
                    onClick={() => setPicker({ kind: "bench" })}
                    className="rounded-xl border border-dashed border-white/20 p-2.5 flex items-center justify-center gap-1.5 text-xs text-slate-500 hover:text-emerald-400 hover:border-emerald-500/40 no-print"
                  >
                    <Plus className="w-3.5 h-3.5" /> Aggiungi riserva
                  </button>
                )}
              </div>

              <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide mb-2">Punti di forza e di debolezza del modulo</p>
              <StrengthsWeaknesses formation={formation} />
              {formation.note && (
                <p className="text-xs text-amber-300/90 bg-amber-500/10 border border-amber-500/20 rounded-xl px-3 py-2.5 mt-4 italic">
                  {formation.note}
                </p>
              )}
              <PositionLegend formation={formation} />
            </div>
          </div>
        </div>
      )}

      <PlayerPickerModal
        open={!!picker}
        onClose={() => setPicker(null)}
        title={picker?.kind === "bench" ? "Aggiungi riserva" : "Scegli giocatore"}
        players={players}
        excludedIds={assignedIds}
        currentId={picker?.currentId}
        suggestedRole={picker?.suggestedRole}
        onSelect={(playerId) => {
          if (picker?.kind === "bench") addToBench(playerId);
          else if (picker?.kind === "position") assignToPosition(picker.positionId, playerId);
        }}
        onClear={
          picker?.kind === "position" && picker?.currentId
            ? () => {
                clearPosition(picker.positionId);
                setPicker(null);
              }
            : null
        }
        onUpdateStatus={updatePlayerStatus}
      />

      <Modal open={showNewFormation} onClose={() => setShowNewFormation(false)} title="Crea nuovo modulo" wide>
        {showNewFormation && (
          <NewFormationWizard onCancel={() => setShowNewFormation(false)} onSave={saveCustomFormation} />
        )}
      </Modal>
    </div>
  );
}

/* ============================================================
   SEZIONE CAMPIONATO
   ============================================================ */

function ChampionshipSection({ season, updateSeason, showToast }) {
  const championship = season.championship || { fase1: null, fase2: null, faseFinale: null };
  const [activeTab, setActiveTab] = useState(championship.fase1 ? "fase1" : "fase1");
  const [showTeamForm, setShowTeamForm] = useState(false);
  const [editingTeam, setEditingTeam] = useState(null);
  const [showMatchForm, setShowMatchForm] = useState(false);
  const [editingMatch, setEditingMatch] = useState(null);
  const [confirmDeleteTeamId, setConfirmDeleteTeamId] = useState(null);
  const [confirmDeleteMatchId, setConfirmDeleteMatchId] = useState(null);
  const [confirmClosePhase, setConfirmClosePhase] = useState(false);
  const [confirmReopenPhase, setConfirmReopenPhase] = useState(false);

  const phase = championship[activeTab];

  useEffect(() => {
    setConfirmClosePhase(false);
    setConfirmReopenPhase(false);
  }, [activeTab]);
  const meta = CHAMPIONSHIP_PHASE_META[activeTab];
  const showStandings = activeTab === "fase1" || activeTab === "fase2";

  // Una nuova fase si può creare solo se la fase "precedente" rilevante è stata chiusa
  function priorPhaseClosed(key) {
    if (key === "fase1") return true;
    if (key === "fase2") return !!championship.fase1?.closed;
    if (key === "faseFinale") {
      if (championship.fase2) return !!championship.fase2.closed;
      return !!championship.fase1?.closed;
    }
    return false;
  }

  function createPhase(key) {
    updateSeason((s) => ({
      championship: {
        ...(s.championship || { fase1: null, fase2: null, faseFinale: null }),
        [key]: emptyChampionshipPhase(s),
      },
    }));
    setActiveTab(key);
    showToast(`${CHAMPIONSHIP_PHASE_META[key].label} creata`);
  }

  function updatePhase(updater) {
    updateSeason((s) => {
      const current = (s.championship || {})[activeTab];
      if (!current) return {};
      return {
        championship: { ...(s.championship || {}), [activeTab]: { ...current, ...updater(current) } },
      };
    });
  }

  function saveTeam(data) {
    if (editingTeam) {
      updatePhase((p) => ({ teams: p.teams.map((t) => (t.id === editingTeam.id ? { ...t, ...data } : t)) }));
      showToast("Squadra aggiornata");
    } else {
      updatePhase((p) => ({ teams: [...p.teams, { ...data, id: uid("cteam") }] }));
      showToast("Squadra aggiunta");
    }
    setShowTeamForm(false);
    setEditingTeam(null);
  }

  function deleteTeam(teamId) {
    updatePhase((p) => ({
      teams: p.teams.filter((t) => t.id !== teamId),
      matches: p.matches.filter((m) => m.homeTeamId !== teamId && m.awayTeamId !== teamId),
    }));
    setConfirmDeleteTeamId(null);
    showToast("Squadra rimossa");
  }

  function addOrUpdateMatch(match) {
    updatePhase((p) => {
      const exists = p.matches.some((m) => m.id === match.id);
      return {
        matches: exists ? p.matches.map((m) => (m.id === match.id ? match : m)) : [...p.matches, { ...match, id: uid("cmatch") }],
      };
    });
    setShowMatchForm(false);
    setEditingMatch(null);
  }

  function deleteMatch(matchId) {
    updatePhase((p) => ({ matches: p.matches.filter((m) => m.id !== matchId) }));
    setConfirmDeleteMatchId(null);
  }

  function closePhase() {
    updatePhase(() => ({ closed: true }));
    setConfirmClosePhase(false);
    showToast(`${meta.label} chiusa`);
  }

  function hasSubsequentPhase(key) {
    if (key === "fase1") return !!championship.fase2 || !!championship.faseFinale;
    if (key === "fase2") return !!championship.faseFinale;
    return false;
  }

  function requestReopenPhase() {
    if (hasSubsequentPhase(activeTab)) {
      setConfirmReopenPhase(true);
    } else {
      reopenPhase();
    }
  }

  function reopenPhase() {
    updateSeason((s) => {
      const champ = { ...(s.championship || { fase1: null, fase2: null, faseFinale: null }) };
      if (activeTab === "fase1") {
        champ.fase2 = null;
        champ.faseFinale = null;
      } else if (activeTab === "fase2") {
        champ.faseFinale = null;
      }
      champ[activeTab] = { ...champ[activeTab], closed: false };
      return { championship: champ };
    });
    setConfirmReopenPhase(false);
    showToast(`${meta.label} riaperta`);
  }

  const standings = phase && showStandings ? computeStandings(phase.teams, phase.matches) : [];
  const sortedMatches = phase ? [...phase.matches].sort((a, b) => new Date(b.date) - new Date(a.date)) : [];

  function teamById(id) {
    return phase?.teams.find((t) => t.id === id);
  }

  return (
    <div>
      <SectionTitle
        eyebrow="Torneo"
        title={`Campionato — ${meta.label}`}
        icon={Award}
        action={
          phase &&
          (phase.closed ? (
            <div className="flex flex-col items-end gap-2">
              <Badge className="bg-slate-500/15 text-slate-400 border-slate-500/30 text-sm">
                <CheckCircle2 className="w-3.5 h-3.5" /> Fase chiusa
              </Badge>
              {confirmReopenPhase ? (
                <div className="text-right max-w-xs">
                  <p className="text-[11px] text-amber-300 mb-1.5">
                    Sono state create fasi successive: riaprendo questa fase verranno eliminate definitivamente. Confermi?
                  </p>
                  <div className="flex gap-2 justify-end">
                    <Button variant="secondary" onClick={() => setConfirmReopenPhase(false)}>Annulla</Button>
                    <Button variant="danger" onClick={reopenPhase}>Conferma eliminazione e riapri</Button>
                  </div>
                </div>
              ) : (
                <button
                  onClick={requestReopenPhase}
                  className="rounded-xl text-white px-4 py-2.5 text-sm font-semibold flex items-center gap-2 transition-opacity hover:opacity-90"
                  style={{ backgroundColor: "#d97706" }}
                >
                  <RefreshCw className="w-4 h-4" /> Riapertura Fase
                </button>
              )}
            </div>
          ) : confirmClosePhase ? (
            <div className="flex items-center gap-2">
              <span className="text-xs text-rose-300">Confermi la chiusura?</span>
              <Button variant="secondary" onClick={() => setConfirmClosePhase(false)}>Annulla</Button>
              <Button variant="danger" onClick={closePhase}>Conferma</Button>
            </div>
          ) : (
            <Button variant="danger" onClick={() => setConfirmClosePhase(true)}>
              <XCircle className="w-4 h-4" /> Chiusura Fase
            </Button>
          ))
        }
      />

      <div className="flex items-center justify-between gap-3 mb-5 flex-wrap">
        <div className="flex gap-2 flex-wrap">
          {Object.keys(CHAMPIONSHIP_PHASE_META).map((key) => {
            const exists = !!championship[key];
            const canCreate = priorPhaseClosed(key);
            return (
              <button
                key={key}
                disabled={!exists && !canCreate}
                onClick={() => setActiveTab(key)}
                className={`rounded-full px-4 py-2 text-xs font-semibold border transition-colors disabled:opacity-30 disabled:cursor-not-allowed ${
                  activeTab === key ? "bg-emerald-500 text-slate-950 border-emerald-500" : "border-white/10 text-slate-400"
                }`}
              >
                {CHAMPIONSHIP_PHASE_META[key].label}
                {!exists && canCreate && " (da creare)"}
                {championship[key]?.closed && " ✓"}
              </button>
            );
          })}
        </div>
        <SectionResetButton
          label="Azzera Campionato"
          confirmText="Eliminare tutte le fasi del campionato (squadre, partite, classifiche)?"
          onConfirm={() => {
            updateSeason(() => ({ championship: { fase1: null, fase2: null, faseFinale: null } }));
            showToast("Campionato azzerato");
          }}
        />
      </div>

      {!phase ? (
        <Card className="p-6 text-center">
          <Award className="w-8 h-8 text-slate-600 mx-auto mb-3" />
          <p className="text-sm text-slate-300 mb-1">
            {meta.label} non ancora creata{meta.period ? ` (${meta.period})` : ""}.
          </p>
          <p className="text-xs text-slate-500 mb-4">
            {!priorPhaseClosed(activeTab)
              ? "Devi prima chiudere la fase precedente (pulsante rosso \"Chiusura Fase\") prima di poterne creare una nuova."
              : activeTab === "fase2"
              ? "Puoi creare la Fase 2, oppure passare direttamente alla Fase Finale se il campionato non la prevede."
              : activeTab === "faseFinale"
              ? "Puoi creare la Fase Finale in qualsiasi momento, anche saltando la Fase 2, purché la fase precedente sia chiusa."
              : "Inizia da qui: la tua squadra verrà inserita automaticamente."}
          </p>
          {priorPhaseClosed(activeTab) && (
            <Button onClick={() => createPhase(activeTab)}>
              <Plus className="w-4 h-4" /> Crea {meta.label}
            </Button>
          )}
        </Card>
      ) : (
        <div>
          {showStandings && (
            <Card className="p-4 mb-5 overflow-x-auto">
              <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide mb-3">Classifica</p>
              {standings.length === 0 ? (
                <p className="text-sm text-slate-500">Aggiungi squadre e risultati per generare la classifica.</p>
              ) : (
                <table className="w-full text-sm">
                  <thead>
                    <tr className="text-left text-[11px] uppercase tracking-wide text-slate-500">
                      <th className="py-1.5 pr-2">#</th>
                      <th className="py-1.5 pr-2">Squadra</th>
                      <th className="py-1.5 px-1.5 text-center">PG</th>
                      <th className="py-1.5 px-1.5 text-center">V</th>
                      <th className="py-1.5 px-1.5 text-center">N</th>
                      <th className="py-1.5 px-1.5 text-center">P</th>
                      <th className="py-1.5 px-1.5 text-center">GF</th>
                      <th className="py-1.5 px-1.5 text-center">GS</th>
                      <th className="py-1.5 pl-1.5 text-center">Pt</th>
                    </tr>
                  </thead>
                  <tbody>
                    {standings.map((t, i) => (
                      <tr key={t.id} className={`border-t border-white/5 ${t.isUs ? "bg-emerald-500/10" : ""}`}>
                        <td className="py-1.5 pr-2 text-slate-500">{i + 1}</td>
                        <td className="py-1.5 pr-2">
                          <div className="flex items-center gap-2">
                            <ColorPair primary={t.colorPrimary} secondary={t.colorSecondary} size={14} />
                            <span className={t.isUs ? "font-bold text-emerald-400" : "text-slate-200"}>{t.name}</span>
                          </div>
                        </td>
                        <td className="py-1.5 px-1.5 text-center text-slate-400">{t.played}</td>
                        <td className="py-1.5 px-1.5 text-center text-slate-400">{t.w}</td>
                        <td className="py-1.5 px-1.5 text-center text-slate-400">{t.d}</td>
                        <td className="py-1.5 px-1.5 text-center text-slate-400">{t.l}</td>
                        <td className="py-1.5 px-1.5 text-center text-slate-400">{t.gf}</td>
                        <td className="py-1.5 px-1.5 text-center text-slate-400">{t.ga}</td>
                        <td className="py-1.5 pl-1.5 text-center font-bold text-slate-100">{t.pts}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </Card>
          )}

          <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
            <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide">Squadre partecipanti ({phase.teams.length})</p>
            <Button
              variant="secondary"
              onClick={() => {
                setEditingTeam(null);
                setShowTeamForm(true);
              }}
              disabled={phase.matches.length > 0}
              title={phase.matches.length > 0 ? "Non puoi più aggiungere squadre dopo aver inserito la prima partita" : ""}
            >
              <Plus className="w-4 h-4" /> Aggiungi Squadra
            </Button>
          </div>
          {phase.matches.length > 0 && (
            <p className="text-[11px] text-slate-500 -mt-2 mb-3">
              L'elenco squadre è bloccato: è già stata inserita almeno una partita. Puoi comunque rinominare o cambiare i colori.
            </p>
          )}
          <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-2 mb-6">
            {phase.teams.map((t) => (
              <div key={t.id} className={`flex items-center justify-between gap-2 rounded-xl border p-2.5 ${t.isUs ? "border-emerald-500/40 bg-emerald-500/5" : "border-white/10"}`}>
                <div className="flex items-center gap-2 min-w-0">
                  <ColorPair primary={t.colorPrimary} secondary={t.colorSecondary} size={16} />
                  <span className={`text-sm truncate ${t.isUs ? "font-bold text-emerald-400" : "text-slate-200"}`}>{t.name}</span>
                  {t.isUs && <Badge className="bg-emerald-500/15 text-emerald-400 border-emerald-500/30 shrink-0">Noi</Badge>}
                </div>
                <div className="flex items-center gap-1 shrink-0">
                  <button
                    onClick={() => {
                      setEditingTeam(t);
                      setShowTeamForm(true);
                    }}
                    className="text-slate-400 hover:text-slate-200 p-1"
                    title="Modifica nome e colori"
                  >
                    <Edit2 className="w-3.5 h-3.5" />
                  </button>
                  {!t.isUs && (
                    <button onClick={() => setConfirmDeleteTeamId(t.id)} className="text-rose-400 p-1">
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  )}
                </div>
              </div>
            ))}
          </div>
          {confirmDeleteTeamId && (
            <div className="rounded-xl border border-rose-500/30 bg-rose-500/10 p-3 mb-5 flex items-center justify-between gap-2">
              <p className="text-sm text-rose-300">Rimuovere la squadra? Le partite collegate verranno eliminate.</p>
              <div className="flex gap-2 shrink-0">
                <Button variant="secondary" onClick={() => setConfirmDeleteTeamId(null)}>Annulla</Button>
                <Button variant="danger" onClick={() => deleteTeam(confirmDeleteTeamId)}>Elimina</Button>
              </div>
            </div>
          )}

          <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
            <p className="text-xs font-semibold text-slate-400 uppercase tracking-wide">Partite ({phase.matches.length})</p>
            <Button
              onClick={() => {
                setEditingMatch(null);
                setShowMatchForm(true);
              }}
              disabled={phase.teams.length < 2 || phase.closed}
              title={phase.closed ? "Fase chiusa: non è possibile aggiungere nuove partite" : ""}
            >
              <Plus className="w-4 h-4" /> Aggiungi Partita
            </Button>
          </div>
          {phase.closed && (
            <p className="text-[11px] text-slate-500 -mt-2 mb-3">
              Questa fase è chiusa: puoi ancora correggere i risultati delle partite esistenti, ma non aggiungerne di nuove.
            </p>
          )}

          {sortedMatches.length === 0 ? (
            <EmptyState icon={Trophy} text={phase.teams.length < 2 ? "Aggiungi almeno 2 squadre per poter inserire le partite." : "Nessuna partita inserita ancora."} />
          ) : (
            <div className="space-y-2">
              {sortedMatches.map((m) => {
                const home = teamById(m.homeTeamId);
                const away = teamById(m.awayTeamId);
                const usInvolved = home?.isUs || away?.isUs;
                return (
                  <button
                    key={m.id}
                    onClick={() => {
                      setEditingMatch(m);
                      setShowMatchForm(true);
                    }}
                    className="w-full text-left"
                  >
                    <Card className={`p-3 hover:border-emerald-500/40 transition-colors ${usInvolved ? "border-emerald-500/20" : ""}`}>
                      <div className="flex items-center justify-between gap-3 flex-wrap">
                        <div className="flex items-center gap-2 min-w-0">
                          <ColorPair primary={home?.colorPrimary} secondary={home?.colorSecondary} size={14} />
                          <span className={`text-sm truncate ${home?.isUs ? "font-bold text-emerald-400" : "text-slate-200"}`}>{home?.name || "?"}</span>
                          <span className="text-slate-600 text-xs">vs</span>
                          <span className={`text-sm truncate ${away?.isUs ? "font-bold text-emerald-400" : "text-slate-200"}`}>{away?.name || "?"}</span>
                          <ColorPair primary={away?.colorPrimary} secondary={away?.colorSecondary} size={14} />
                        </div>
                        <div className="flex items-center gap-3 shrink-0">
                          <span className="text-[11px] text-slate-500">{formatDate(m.date)}</span>
                          {m.played ? (
                            <span className="text-base font-extrabold text-slate-100">{m.homeGoals} - {m.awayGoals}</span>
                          ) : (
                            <Badge className="bg-sky-500/15 text-sky-400 border-sky-500/30">Da giocare</Badge>
                          )}
                        </div>
                      </div>
                    </Card>
                  </button>
                );
              })}
            </div>
          )}
        </div>
      )}

      <Modal
        open={showTeamForm}
        onClose={() => { setShowTeamForm(false); setEditingTeam(null); }}
        title={editingTeam ? "Modifica Squadra" : "Aggiungi Squadra"}
      >
        <ChampionshipTeamForm
          initial={editingTeam}
          onSubmit={saveTeam}
          onCancel={() => { setShowTeamForm(false); setEditingTeam(null); }}
        />
      </Modal>

      <Modal open={showMatchForm} onClose={() => { setShowMatchForm(false); setEditingMatch(null); }} title={editingMatch ? "Modifica Partita" : "Aggiungi Partita"}>
        <ChampionshipMatchForm
          teams={phase?.teams || []}
          initial={editingMatch}
          realMatches={season.matches || []}
          onSubmit={addOrUpdateMatch}
          onCancel={() => { setShowMatchForm(false); setEditingMatch(null); }}
          onDelete={editingMatch ? () => deleteMatch(editingMatch.id) : null}
        />
      </Modal>
    </div>
  );
}

function ChampionshipTeamForm({ onSubmit, onCancel, initial }) {
  const [name, setName] = useState(initial?.name || "");
  const [colorPrimary, setColorPrimary] = useState(initial?.colorPrimary || "#dc2626");
  const [colorSecondary, setColorSecondary] = useState(initial?.colorSecondary || "#0f172a");
  const isEditing = !!initial;
  return (
    <div>
      <Field label="Nome squadra">
        <input className={inputClass} value={name} onChange={(e) => setName(e.target.value)} placeholder="Es. Sestrese" />
      </Field>
      <Field label="Colore primario">
        <ColorSwatchPicker value={colorPrimary} onChange={setColorPrimary} />
      </Field>
      <Field label="Colore secondario">
        <ColorSwatchPicker value={colorSecondary} onChange={setColorSecondary} />
      </Field>
      <div className="flex justify-end gap-2 mt-4">
        <Button variant="secondary" onClick={onCancel}>Annulla</Button>
        <Button onClick={() => name.trim() && onSubmit({ name: name.trim(), colorPrimary, colorSecondary, isUs: initial?.isUs || false })}>
          <Save className="w-4 h-4" /> {isEditing ? "Salva Modifiche" : "Aggiungi"}
        </Button>
      </div>
    </div>
  );
}

function ChampionshipMatchForm({ teams, initial, onSubmit, onCancel, onDelete, realMatches }) {
  const [form, setForm] = useState(
    initial || {
      homeTeamId: teams[0]?.id || "",
      awayTeamId: teams[1]?.id || "",
      date: todayISO(),
      played: false,
      homeGoals: 0,
      awayGoals: 0,
    }
  );
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [importId, setImportId] = useState("");
  const [importError, setImportError] = useState("");
  const invalid = form.homeTeamId && form.awayTeamId && form.homeTeamId === form.awayTeamId;

  const usTeam = teams.find((t) => t.isUs);
  const usIsHome = usTeam && form.homeTeamId === usTeam.id;
  const usIsAway = usTeam && form.awayTeamId === usTeam.id;
  const usInvolved = usIsHome || usIsAway;
  // Solo le partite di campionato hanno senso da importare in una fase di campionato
  const playedRealMatches = (realMatches || []).filter((m) => m.status === "Disputata" && m.result && m.matchType === "Campionato");

  function applyImport(matchId) {
    setImportId(matchId);
    setImportError("");
    const rm = playedRealMatches.find((m) => m.id === matchId);
    if (!rm) return;

    // Cerca tra le squadre già inserite in questa fase quella con lo stesso nome dell'avversario reale
    const matchedTeam = teams.find((t) => !t.isUs && t.name.trim().toLowerCase() === rm.opponent.trim().toLowerCase());
    if (!matchedTeam) {
      setImportError(`La squadra "${rm.opponent}" non è tra quelle inserite in questa fase. Aggiungila prima di importare questo risultato.`);
      return;
    }

    // "golFor" nella partita reale è sempre riferito alla nostra squadra
    if (usIsHome) {
      setForm({ ...form, awayTeamId: matchedTeam.id, played: true, homeGoals: rm.result.golFor, awayGoals: rm.result.golAgainst });
    } else if (usIsAway) {
      setForm({ ...form, homeTeamId: matchedTeam.id, played: true, awayGoals: rm.result.golFor, homeGoals: rm.result.golAgainst });
    }
  }

  return (
    <div>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Squadra Casa">
          <select className={inputClass} value={form.homeTeamId} onChange={(e) => { setForm({ ...form, homeTeamId: e.target.value }); setImportId(""); }}>
            {teams.map((t) => (
              <option key={t.id} value={t.id}>{t.name}{t.isUs ? " (Noi)" : ""}</option>
            ))}
          </select>
        </Field>
        <Field label="Squadra Trasferta">
          <select className={inputClass} value={form.awayTeamId} onChange={(e) => { setForm({ ...form, awayTeamId: e.target.value }); setImportId(""); }}>
            {teams.map((t) => (
              <option key={t.id} value={t.id}>{t.name}{t.isUs ? " (Noi)" : ""}</option>
            ))}
          </select>
        </Field>
      </div>
      {invalid && <p className="text-[11px] text-rose-400 mb-2">Le due squadre devono essere diverse.</p>}

      {usInvolved && playedRealMatches.length > 0 && (
        <Field label="Importa risultato da una partita di campionato già disputata (facoltativo)">
          <select className={inputClass} value={importId} onChange={(e) => e.target.value && applyImport(e.target.value)}>
            <option value="">— Inserisci il risultato manualmente —</option>
            {playedRealMatches.map((m) => (
              <option key={m.id} value={m.id}>
                vs {m.opponent} · {formatDate(m.date)} · {m.result.golFor}-{m.result.golAgainst}
              </option>
            ))}
          </select>
          {importError && (
            <p className="text-[11px] text-rose-400 mt-1.5 flex items-center gap-1">
              <AlertCircle className="w-3.5 h-3.5 shrink-0" /> {importError}
            </p>
          )}
        </Field>
      )}

      <div className="flex items-center gap-2 mb-2">
        <input
          type="checkbox"
          id="played-checkbox"
          checked={form.played}
          onChange={(e) => setForm({ ...form, played: e.target.checked })}
          className="w-4 h-4 accent-emerald-500"
        />
        <label htmlFor="played-checkbox" className="text-sm text-slate-300">Partita disputata (inserisci il risultato)</label>
      </div>

      {form.played && (
        <div className="flex items-center justify-center gap-4 my-4">
          <input
            type="number"
            min="0"
            className="w-16 rounded-xl border border-white/10 bg-slate-950/60 text-center text-2xl font-extrabold py-2 text-slate-100"
            value={form.homeGoals}
            onChange={(e) => setForm({ ...form, homeGoals: Number(e.target.value) })}
          />
          <span className="text-slate-600 text-xl">-</span>
          <input
            type="number"
            min="0"
            className="w-16 rounded-xl border border-white/10 bg-slate-950/60 text-center text-2xl font-extrabold py-2 text-slate-100"
            value={form.awayGoals}
            onChange={(e) => setForm({ ...form, awayGoals: Number(e.target.value) })}
          />
        </div>
      )}

      {confirmDelete ? (
        <div className="rounded-xl border border-rose-500/30 bg-rose-500/10 p-3 mb-3 flex items-center justify-between gap-2">
          <p className="text-sm text-rose-300">Eliminare questa partita?</p>
          <div className="flex gap-2 shrink-0">
            <Button variant="secondary" onClick={() => setConfirmDelete(false)}>Annulla</Button>
            <Button variant="danger" onClick={onDelete}>Elimina</Button>
          </div>
        </div>
      ) : (
        <div className="flex justify-between gap-2 mt-4">
          {onDelete ? (
            <Button variant="danger" onClick={() => setConfirmDelete(true)}>
              <Trash2 className="w-4 h-4" /> Elimina
            </Button>
          ) : (
            <span />
          )}
          <div className="flex gap-2">
            <Button variant="secondary" onClick={onCancel}>Annulla</Button>
            <Button disabled={invalid || !form.homeTeamId || !form.awayTeamId} onClick={() => onSubmit(form)}>
              <Save className="w-4 h-4" /> Salva
            </Button>
          </div>
        </div>
      )}
    </div>
  );
}

/* ============================================================
   SEZIONE CONFIGURAZIONI (tabelle personalizzabili)
   ============================================================ */

const CONFIG_TABLES = [
  { key: "roles", label: "Ruoli", hint: "Usato per Ruolo Principale e Ruolo Alternativo dei giocatori" },
  { key: "positions", label: "Posizione in Campo", hint: "" },
  { key: "medicalStatuses", label: "Stato Medico", hint: "" },
  { key: "exerciseTypes", label: "Tipologia Esercizi", hint: "" },
  { key: "categories", label: "Categorie", hint: "Fasce d'età/categoria del club, usate per taggare gli esercizi" },
];

function ConfigurationsSection({ library, updateLibrary, showToast }) {
  const config = library.config || defaultConfig();
  const [activeTable, setActiveTable] = useState("roles");
  const meta = CONFIG_TABLES.find((t) => t.key === activeTable);
  const items = config[activeTable] || [];

  function setItems(newItems) {
    updateLibrary((lib) => ({ config: { ...(lib.config || defaultConfig()), [activeTable]: newItems } }));
  }

  return (
    <div className="max-w-2xl">
      <SectionTitle eyebrow="Personalizzazione" title="Configurazioni" icon={SlidersHorizontal} />
      <p className="text-sm text-slate-400 mb-5">
        Gestisci le voci preimpostate usate nei menu a tendina dell'app: aggiungi, rinomina o elimina liberamente.
      </p>

      <div className="flex gap-2 mb-5 flex-wrap">
        {CONFIG_TABLES.map((t) => (
          <button
            key={t.key}
            onClick={() => setActiveTable(t.key)}
            className={`rounded-full px-4 py-2 text-xs font-semibold border transition-colors ${
              activeTable === t.key ? "bg-emerald-500 text-slate-950 border-emerald-500" : "border-white/10 text-slate-400"
            }`}
          >
            {t.label}
          </button>
        ))}
      </div>

      <Card className="p-5">
        <p className="text-sm font-bold text-slate-100 mb-1">{meta.label}</p>
        {meta.hint && <p className="text-xs text-slate-500 mb-4">{meta.hint}</p>}
        <EditableListManager items={items} onChange={setItems} showToast={showToast} />
      </Card>
    </div>
  );
}

function EditableListManager({ items, onChange, showToast }) {
  const [newValue, setNewValue] = useState("");
  const [editingIdx, setEditingIdx] = useState(null);
  const [editValue, setEditValue] = useState("");
  const [confirmDeleteIdx, setConfirmDeleteIdx] = useState(null);

  function addItem() {
    const v = newValue.trim();
    if (!v) return;
    if (items.some((i) => i.toLowerCase() === v.toLowerCase())) {
      showToast("Voce già presente nell'elenco", "error");
      return;
    }
    onChange([...items, v]);
    setNewValue("");
  }

  function saveEdit(idx) {
    const v = editValue.trim();
    if (!v) return;
    const updated = [...items];
    updated[idx] = v;
    onChange(updated);
    setEditingIdx(null);
  }

  function deleteItem(idx) {
    onChange(items.filter((_, i) => i !== idx));
    setConfirmDeleteIdx(null);
  }

  return (
    <div>
      <div className="flex gap-2 mb-4">
        <input
          className={inputClass}
          value={newValue}
          onChange={(e) => setNewValue(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && addItem()}
          placeholder="Nuova voce..."
        />
        <Button onClick={addItem}>
          <Plus className="w-4 h-4" /> Aggiungi
        </Button>
      </div>

      {items.length === 0 ? (
        <EmptyState icon={SlidersHorizontal} text="Nessuna voce presente." />
      ) : (
        <div className="space-y-2">
          {items.map((item, idx) => (
            <div key={idx} className="flex items-center gap-2 rounded-xl border border-white/10 p-2.5">
              {editingIdx === idx ? (
                <>
                  <input
                    className={inputClass}
                    value={editValue}
                    onChange={(e) => setEditValue(e.target.value)}
                    onKeyDown={(e) => e.key === "Enter" && saveEdit(idx)}
                    autoFocus
                  />
                  <Button className="px-3 py-2 text-xs" onClick={() => saveEdit(idx)}>Salva</Button>
                  <Button variant="secondary" className="px-3 py-2 text-xs" onClick={() => setEditingIdx(null)}>Annulla</Button>
                </>
              ) : confirmDeleteIdx === idx ? (
                <>
                  <span className="flex-1 text-sm text-rose-300">Eliminare "{item}"?</span>
                  <Button variant="danger" className="px-3 py-2 text-xs" onClick={() => deleteItem(idx)}>Elimina</Button>
                  <Button variant="secondary" className="px-3 py-2 text-xs" onClick={() => setConfirmDeleteIdx(null)}>Annulla</Button>
                </>
              ) : (
                <>
                  <span className="flex-1 text-sm text-slate-200">{item}</span>
                  <button
                    onClick={() => {
                      setEditingIdx(idx);
                      setEditValue(item);
                    }}
                    className="rounded-lg p-1.5 hover:bg-white/10 text-slate-400"
                  >
                    <Edit2 className="w-3.5 h-3.5" />
                  </button>
                  <button onClick={() => setConfirmDeleteIdx(idx)} className="rounded-lg p-1.5 hover:bg-rose-500/10 text-rose-400">
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                </>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

/* ============================================================
   SEZIONE DOSSIER
   ============================================================ */

function fileSizeLabel(bytes) {
  if (!bytes) return "";
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function DossierSection({ library, updateLibrary, showToast }) {
  const dossier = library.dossier || [];
  const [search, setSearch] = useState("");
  const [categoryFilter, setCategoryFilter] = useState("Tutti");
  const [pendingFile, setPendingFile] = useState(null); // { fileName, fileType, fileSize, dataUrl }
  const [showUploadForm, setShowUploadForm] = useState(false);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [previewDoc, setPreviewDoc] = useState(null);
  const fileInputRef = React.useRef(null);

  function handleFileSelected(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    if (file.size > 4 * 1024 * 1024) {
      showToast("File troppo grande (max consigliato 4 MB): il salvataggio potrebbe fallire", "error");
    }
    const reader = new FileReader();
    reader.onload = () => {
      setPendingFile({
        fileName: file.name,
        fileType: file.type || "application/octet-stream",
        fileSize: file.size,
        dataUrl: reader.result,
      });
      setShowUploadForm(true);
    };
    reader.onerror = () => showToast("Impossibile leggere il file", "error");
    reader.readAsDataURL(file);
    e.target.value = "";
  }

  function saveDocument({ title, category }) {
    if (!pendingFile) return;
    updateLibrary((lib) => ({
      dossier: [
        ...(lib.dossier || []),
        {
          id: uid("doc"),
          title: title.trim(),
          category,
          fileName: pendingFile.fileName,
          fileType: pendingFile.fileType,
          fileSize: pendingFile.fileSize,
          dataUrl: pendingFile.dataUrl,
          uploadedAt: new Date().toISOString(),
        },
      ],
    }));
    setShowUploadForm(false);
    setPendingFile(null);
    showToast("Documento caricato nel Dossier");
  }

  function deleteDocument(id) {
    const removed = dossier.find((d) => d.id === id);
    updateLibrary((lib) => ({ dossier: (lib.dossier || []).filter((d) => d.id !== id) }));
    setConfirmDeleteId(null);
    showToast(
      `"${removed?.title || "Documento"}" eliminato`,
      "success",
      removed
        ? { label: "Annulla", onClick: () => updateLibrary((lib) => ({ dossier: [...(lib.dossier || []), removed] })) }
        : null
    );
  }

  function openDocument(doc) {
    const a = document.createElement("a");
    a.href = doc.dataUrl;
    a.download = doc.fileName;
    a.target = "_blank";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  const filtered = dossier.filter((d) => {
    const q = search.toLowerCase();
    const matchesSearch = !q || d.title.toLowerCase().includes(q) || d.fileName.toLowerCase().includes(q);
    const matchesCategory = categoryFilter === "Tutti" || d.category === categoryFilter;
    return matchesSearch && matchesCategory;
  });

  return (
    <div>
      <SectionTitle
        eyebrow="Archivio"
        title={`Dossier (${dossier.length})`}
        icon={FileText}
        action={
          <Button onClick={() => fileInputRef.current?.click()}>
            <Plus className="w-4 h-4" /> Carica Documento
          </Button>
        }
      />
      <input ref={fileInputRef} type="file" className="hidden" onChange={handleFileSelected} />

      <div className="flex flex-col sm:flex-row gap-3 mb-3">
        <div className="relative flex-1">
          <Search className="w-4 h-4 text-slate-500 absolute left-3 top-1/2 -translate-y-1/2" />
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Cerca documento per titolo o nome file..."
            className={inputClass + " pl-9"}
          />
        </div>
      </div>

      <div className="flex items-center gap-2 mb-5 overflow-x-auto pb-1">
        {["Tutti", ...DOSSIER_CATEGORIES].map((c) => (
          <button
            key={c}
            onClick={() => setCategoryFilter(c)}
            className={`shrink-0 rounded-full px-3.5 py-1.5 text-[11px] font-medium border transition-colors ${
              categoryFilter === c ? "bg-emerald-500 text-slate-950 border-emerald-500" : "border-white/10 text-slate-400 hover:text-slate-200"
            }`}
          >
            {c}
          </button>
        ))}
      </div>

      {filtered.length === 0 ? (
        <EmptyState icon={FileText} text={dossier.length === 0 ? "Nessun documento caricato ancora." : "Nessun documento trovato per questa ricerca."} />
      ) : (
        <div className="space-y-2">
          {filtered.map((d) => (
            <Card key={d.id} className="p-3">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-lg bg-sky-500/15 flex items-center justify-center shrink-0">
                  <FileText className="w-5 h-5 text-sky-400" />
                </div>
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2 flex-wrap">
                    <p className="text-sm font-semibold text-slate-100 truncate">{d.title}</p>
                    <Badge className="bg-white/5 text-slate-300 border-white/10">{d.category}</Badge>
                  </div>
                  <p className="text-[11px] text-slate-500 truncate">
                    {d.fileName} · {fileSizeLabel(d.fileSize)} · {formatDate(d.uploadedAt)}
                  </p>
                </div>
                <div className="flex gap-1 shrink-0">
                  <button onClick={() => setPreviewDoc(d)} className="rounded-lg p-1.5 hover:bg-white/10 text-slate-400" title="Anteprima">
                    <Eye className="w-3.5 h-3.5" />
                  </button>
                  <button
                    onClick={() => shareOrFallback({ dataUrl: d.dataUrl, filename: d.fileName, label: d.title, showToast })}
                    className="rounded-lg p-1.5 hover:bg-white/10 text-slate-400"
                    title="Condividi (es. su WhatsApp)"
                  >
                    <Share2 className="w-3.5 h-3.5" />
                  </button>
                  <button onClick={() => openDocument(d)} className="rounded-lg p-1.5 hover:bg-white/10 text-slate-400" title="Apri / Scarica">
                    <Download className="w-3.5 h-3.5" />
                  </button>
                  <button onClick={() => setConfirmDeleteId(d.id)} className="rounded-lg p-1.5 hover:bg-rose-500/10 text-rose-400" title="Elimina">
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                </div>
              </div>
              {confirmDeleteId === d.id && (
                <div className="mt-2 rounded-lg border border-rose-500/30 bg-rose-500/10 p-2 flex items-center justify-between gap-2">
                  <p className="text-[11px] text-rose-300">Eliminare questo documento?</p>
                  <div className="flex gap-1.5">
                    <Button variant="secondary" className="px-2 py-1 text-xs" onClick={() => setConfirmDeleteId(null)}>Annulla</Button>
                    <Button variant="danger" className="px-2 py-1 text-xs" onClick={() => deleteDocument(d.id)}>Elimina</Button>
                  </div>
                </div>
              )}
            </Card>
          ))}
        </div>
      )}

      <Modal
        open={showUploadForm}
        onClose={() => { setShowUploadForm(false); setPendingFile(null); }}
        title="Carica Documento"
      >
        {pendingFile && (
          <DossierUploadForm
            fileName={pendingFile.fileName}
            fileSize={pendingFile.fileSize}
            onSubmit={saveDocument}
            onCancel={() => { setShowUploadForm(false); setPendingFile(null); }}
          />
        )}
      </Modal>

      {previewDoc && <DossierPreviewModal doc={previewDoc} onClose={() => setPreviewDoc(null)} showToast={showToast} />}
    </div>
  );
}

function DossierUploadForm({ fileName, fileSize, onSubmit, onCancel }) {
  const defaultTitle = fileName.replace(/\.[^/.]+$/, "");
  const [title, setTitle] = useState(defaultTitle);
  const [category, setCategory] = useState(DOSSIER_CATEGORIES[0]);

  return (
    <div>
      <div className="rounded-xl border border-white/10 p-3 mb-4 flex items-center gap-3">
        <FileText className="w-5 h-5 text-sky-400 shrink-0" />
        <div className="min-w-0">
          <p className="text-sm text-slate-200 truncate">{fileName}</p>
          <p className="text-[11px] text-slate-500">{fileSizeLabel(fileSize)}</p>
        </div>
      </div>
      <Field label="Titolo">
        <input className={inputClass} value={title} onChange={(e) => setTitle(e.target.value)} placeholder="Es. Regolamento campionato Esordienti" />
      </Field>
      <Field label="Categoria">
        <select className={inputClass} value={category} onChange={(e) => setCategory(e.target.value)}>
          {DOSSIER_CATEGORIES.map((c) => (
            <option key={c} value={c}>{c}</option>
          ))}
        </select>
      </Field>
      <div className="flex justify-end gap-2 mt-4">
        <Button variant="secondary" onClick={onCancel}>Annulla</Button>
        <Button onClick={() => title.trim() && onSubmit({ title, category })}>
          <Save className="w-4 h-4" /> Salva nel Dossier
        </Button>
      </div>
    </div>
  );
}

/* ============================================================
   SEZIONE EXPORT / IMPORT
   ============================================================ */

function ExportSection({ seasons, activeSeason, setSeasons, setActiveSeasonId, library, setLibrary, showToast, lastSavedAt, onExported }) {
  const fileInputRef = React.useRef(null);
  const [importPreview, setImportPreview] = useState(null); // { data, fileName, isOlder }

  function exportFullJSON() {
    const payload = JSON.stringify(
      { seasons, activeSeasonId: activeSeason?.id, library, exportedAt: new Date().toISOString() },
      null,
      2
    );
    downloadBlob(payload, `backup-${activeSeason?.name || "stagione"}_${formatBackupDateSuffix()}.json`, "application/json");
    markExportDone();
    onExported?.();
    showToast("Backup JSON scaricato");
  }

  function handleImportClick() {
    fileInputRef.current?.click();
  }

  function applyImportedData(data) {
    setSeasons(data.seasons);
    setActiveSeasonId(data.activeSeasonId || data.seasons[0]?.id);
    if (data.library) {
      const rawLibrary = {
        ...defaultLibrary(),
        ...data.library,
        config: { ...defaultConfig(), ...(data.library.config || {}) },
      };
      setLibrary(migrateLibraryCategories(migrateFocusExercisesToLibrary(data.seasons, rawLibrary)));
    }
    showToast("Dati importati con successo");
  }

  function handleImportFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = JSON.parse(ev.target.result);
        if (data.seasons && Array.isArray(data.seasons)) {
          // Se il backup risale a prima dell'ultimo salvataggio già presente su questo
          // dispositivo, avvisiamo prima di sovrascrivere: capita facilmente lavorando
          // su più dispositivi (es. PC + cellulare) con backup manuali.
          const backupTime = data.exportedAt ? new Date(data.exportedAt).getTime() : 0;
          const localTime = lastSavedAt ? lastSavedAt.getTime() : 0;
          if (backupTime && localTime && backupTime < localTime) {
            setImportPreview({ data, fileName: file.name, isOlder: true });
          } else {
            applyImportedData(data);
          }
        } else {
          showToast("File non valido: struttura dati non riconosciuta", "error");
        }
      } catch (err) {
        showToast("Errore nella lettura del file JSON", "error");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  }

  function exportPlayersExcel() {
    const players = activeSeason?.players || [];
    if (players.length === 0) return showToast("Nessun giocatore da esportare", "error");

    const anagraficaHeaders = ["Numero", "Nome", "Ruolo principale", "Ruolo alternativo", "Posizione", "Piede preferito", "Altezza (cm)", "Stato medico", "Data di nascita", "Note"];
    const baseHeaders = BASE_STAT_KEYS.map((s) => s.label);
    const mentalHeaders = MENTAL_STAT_KEYS.map((s) => s.label);
    const techHeaders = TECH_TACTIC_STAT_KEYS.map((s) => s.label);
    const gkHeaders = GK_STAT_KEYS.map((s) => s.label);

    const headers = [...anagraficaHeaders, ...baseHeaders, ...mentalHeaders, ...techHeaders, ...gkHeaders];
    const groups = [
      { label: "DATI ANAGRAFICI", span: anagraficaHeaders.length },
      { label: "CARATTERISTICHE BASE", span: baseHeaders.length },
      { label: "STATISTICHE MENTALI", span: mentalHeaders.length },
      { label: "TECNICO/TATTICHE", span: techHeaders.length },
      { label: "PORTIERE", span: gkHeaders.length },
    ];

    const rows = players
      .slice()
      .sort((a, b) => (a.number ?? 999) - (b.number ?? 999))
      .map((p) => {
        const anagrafica = [
          p.number ?? "",
          p.name,
          p.role,
          p.role2 || "",
          p.position || "",
          p.preferredFoot || "",
          p.height ?? "",
          p.medicalStatus,
          p.birthDate || "",
          p.notes || "",
        ];
        const base = BASE_STAT_KEYS.map((s) => p.baseStats?.[s.key] ?? "");
        const mental = MENTAL_STAT_KEYS.map((s) => p.mentalStats?.[s.key] ?? "");
        const tech = TECH_TACTIC_STAT_KEYS.map((s) => p.techTacticStats?.[s.key] ?? "");
        const gk = GK_STAT_KEYS.map((s) => (p.role === "Portiere" ? p.gkStats?.[s.key] ?? "" : ""));
        return [...anagrafica, ...base, ...mental, ...tech, ...gk];
      });

    const colWidths = [
      { wch: 8 }, { wch: 20 }, { wch: 16 }, { wch: 16 }, { wch: 11 }, { wch: 12 }, { wch: 11 }, { wch: 13 }, { wch: 13 }, { wch: 22 },
      ...baseHeaders.map(() => ({ wch: 10 })),
      ...mentalHeaders.map(() => ({ wch: 12 })),
      ...techHeaders.map(() => ({ wch: 14 })),
      ...gkHeaders.map(() => ({ wch: 12 })),
    ];

    const ws = buildProfessionalSheet({
      title: `SCHEDA GIOCATORI — ${activeSeason.teamName || ""}`,
      subtitle: `${activeSeason.leva ? activeSeason.leva + " · " : ""}${activeSeason.name}${activeSeason.teamFormat ? " · Calcio a " + activeSeason.teamFormat : ""}`,
      groups,
      headers,
      rows,
      colWidths,
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Giocatori");
    XLSX.writeFile(wb, `giocatori-${activeSeason.name}.xlsx`);
    showToast("Excel giocatori esportato");
  }

  function exportTrainingsExcel() {
    const trainings = activeSeason?.trainings || [];
    const players = activeSeason?.players || [];
    const focusTecnici = activeSeason?.focusTecnici || [];
    if (trainings.length === 0) return showToast("Nessun allenamento da esportare", "error");

    const detailHeaders = ["Data", "Ora", "Focus tecnico", "Durata totale (min)", "Giocatore", "Stato"];
    const summaryHeaders = ["Data", "Ora", "Focus tecnico", "Presenti", "Assenti", "Giustificati", "Infortunati"];
    const detailRows = [];
    const summaryRows = [];
    trainings.forEach((t) => {
      const linkedFocus = focusTecnici.find((f) => f.id === t.focusTecnicoId);
      const values = players.map((p) => t.attendance?.[p.id] || "Non registrato");
      players.forEach((p) => {
        detailRows.push([t.date, t.time, linkedFocus?.title || t.focus || "", linkedFocus ? totalFocusMinutes(linkedFocus) : "", p.name, t.attendance?.[p.id] || "Non registrato"]);
      });
      summaryRows.push([
        t.date,
        t.time,
        linkedFocus?.title || t.focus || "",
        values.filter((v) => v === "Presente").length,
        values.filter((v) => v === "Assente").length,
        values.filter((v) => v === "Giustificato").length,
        values.filter((v) => v === "Infortunato").length,
      ]);
    });

    const wsSummary = buildProfessionalSheet({
      title: `RIEPILOGO PRESENZE — ${activeSeason.teamName || ""}`,
      subtitle: activeSeason.name,
      groups: [{ label: "ALLENAMENTI", span: summaryHeaders.length }],
      headers: summaryHeaders,
      rows: summaryRows,
      colWidths: [{ wch: 12 }, { wch: 8 }, { wch: 28 }, { wch: 10 }, { wch: 10 }, { wch: 12 }, { wch: 12 }],
    });
    const wsDetail = buildProfessionalSheet({
      title: `DETTAGLIO PRESENZE — ${activeSeason.teamName || ""}`,
      subtitle: activeSeason.name,
      groups: [{ label: "PRESENZE PER GIOCATORE", span: detailHeaders.length }],
      headers: detailHeaders,
      rows: detailRows,
      colWidths: [{ wch: 12 }, { wch: 8 }, { wch: 28 }, { wch: 16 }, { wch: 20 }, { wch: 14 }],
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsSummary, "Riepilogo");
    XLSX.utils.book_append_sheet(wb, wsDetail, "Presenze");
    XLSX.writeFile(wb, `presenze-${activeSeason.name}.xlsx`);
    showToast("Excel presenze esportato");
  }

  function exportMatchesExcel() {
    const matches = activeSeason?.matches || [];
    const players = activeSeason?.players || [];
    if (matches.length === 0) return showToast("Nessuna partita da esportare", "error");

    const headers = [
      "Data", "Ora", "Tipo", "Nome Torneo", "Avversario", "Colore avversario 1", "Colore avversario 2",
      "Casa/Trasferta", "Luogo", "Stato", "Gol Fatti", "Gol Subiti",
      "Marcatori", "Assistman", "Ammoniti", "Espulsi", "Convocati", "Annotazioni mister",
    ];
    const rows = matches.map((m) => [
      m.date,
      m.time,
      m.matchType || "",
      m.matchType === "Torneo" ? m.tournamentName || "" : "",
      m.opponent,
      m.opponentColorPrimary || "",
      m.opponentColorSecondary || "",
      m.homeAway,
      m.venue || "",
      m.status,
      m.result?.golFor ?? "",
      m.result?.golAgainst ?? "",
      (m.scorers || []).map((s) => `${players.find((p) => p.id === s.playerId)?.name || "?"} (${s.goals})`).join(", "),
      (m.assists || []).map((s) => `${players.find((p) => p.id === s.playerId)?.name || "?"} (${s.assists})`).join(", "),
      (m.yellowCards || []).map((id) => players.find((p) => p.id === id)?.name || "?").join(", "),
      (m.redCards || []).map((id) => players.find((p) => p.id === id)?.name || "?").join(", "),
      (m.convocati || []).map((id) => players.find((p) => p.id === id)?.name || "?").join(", "),
      m.coachNotes || "",
    ]);

    const ws = buildProfessionalSheet({
      title: `RIEPILOGO PARTITE — ${activeSeason.teamName || ""}`,
      subtitle: activeSeason.name,
      groups: [
        { label: "DATI PARTITA", span: 10 },
        { label: "RISULTATO", span: 2 },
        { label: "TABELLINO", span: 6 },
      ],
      headers,
      rows,
      colWidths: [
        { wch: 12 }, { wch: 8 }, { wch: 12 }, { wch: 18 }, { wch: 18 }, { wch: 14 }, { wch: 14 },
        { wch: 12 }, { wch: 16 }, { wch: 12 }, { wch: 9 }, { wch: 9 },
        { wch: 30 }, { wch: 30 }, { wch: 20 }, { wch: 20 }, { wch: 30 }, { wch: 30 },
      ],
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Partite");
    XLSX.writeFile(wb, `partite-${activeSeason.name}.xlsx`);
    showToast("Excel partite esportato");
  }

  function exportStatisticsExcel() {
    const players = activeSeason?.players || [];
    const trainings = activeSeason?.trainings || [];
    const matches = activeSeason?.matches || [];
    if (players.length === 0) return showToast("Nessun giocatore da esportare", "error");

    const headers = ["Numero", "Nome", "Ruolo", "Presenze allenamento", "Assenze allenamento", "Convocazioni partita", "Reti", "Assist", "Ammonizioni", "Espulsioni"];
    const rows = players
      .slice()
      .sort((a, b) => (a.number ?? 999) - (b.number ?? 999))
      .map((p) => {
        const s = computePlayerStats(p.id, trainings, matches);
        return [p.number ?? "", p.name, p.role, s.presenze, s.assenze, s.convocazioni, s.reti, s.assist, s.ammonizioni, s.espulsioni];
      });

    const ws = buildProfessionalSheet({
      title: `STATISTICHE STAGIONE — ${activeSeason.teamName || ""}`,
      subtitle: activeSeason.name,
      groups: [
        { label: "GIOCATORE", span: 3 },
        { label: "ALLENAMENTI", span: 2 },
        { label: "PARTITE", span: 5 },
      ],
      headers,
      rows,
      colWidths: [{ wch: 8 }, { wch: 20 }, { wch: 16 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 8 }, { wch: 8 }, { wch: 12 }, { wch: 12 }],
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Statistiche");
    XLSX.writeFile(wb, `statistiche-${activeSeason.name}.xlsx`);
    showToast("Excel statistiche esportato");
  }

  return (
    <div>
      <SectionTitle eyebrow="Backup & Report" title="Esporta Dati" icon={FileSpreadsheet} />

      <div className="grid sm:grid-cols-2 gap-4 mb-6">
        <Card className="p-5">
          <div className="flex items-center gap-2 mb-2">
            <Save className="w-4.5 h-4.5 text-emerald-400" />
            <p className="text-sm font-bold text-slate-100">Backup completo (JSON)</p>
          </div>
          <p className="text-xs text-slate-500 mb-4">
            Scarica tutte le stagioni con giocatori, allenamenti e partite. Usalo come backup o per trasferire i dati su un altro dispositivo.
          </p>
          <div className="flex flex-wrap gap-2">
            <Button onClick={exportFullJSON}><Download className="w-4 h-4" /> Esporta JSON</Button>
            <Button variant="secondary" onClick={handleImportClick}><Upload className="w-4 h-4" /> Importa JSON</Button>
            <input ref={fileInputRef} type="file" accept="application/json" className="hidden" onChange={handleImportFile} />
          </div>
          {importPreview?.isOlder && (
            <div className="mt-3 rounded-xl border border-amber-500/30 bg-amber-500/10 p-3">
              <p className="text-xs text-amber-300 mb-2">
                <AlertCircle className="w-3.5 h-3.5 inline -mt-0.5 mr-1" />
                Il file <strong>{importPreview.fileName}</strong> risale al {formatDateTime(importPreview.data.exportedAt)}, mentre i dati già presenti su questo dispositivo sono più recenti (ultimo salvataggio: {formatDateTime(lastSavedAt?.toISOString())}). Importarlo sovrascriverà i dati più recenti con quelli del backup più vecchio.
              </p>
              <div className="flex gap-2">
                <Button variant="secondary" className="px-2.5 py-1 text-xs" onClick={() => setImportPreview(null)}>Annulla</Button>
                <Button
                  variant="danger"
                  className="px-2.5 py-1 text-xs"
                  onClick={() => {
                    applyImportedData(importPreview.data);
                    setImportPreview(null);
                  }}
                >
                  Importa comunque
                </Button>
              </div>
            </div>
          )}
        </Card>

        <Card className="p-5">
          <div className="flex items-center gap-2 mb-2">
            <FileSpreadsheet className="w-4.5 h-4.5 text-emerald-400" />
            <p className="text-sm font-bold text-slate-100">Report Excel — {activeSeason?.name}</p>
          </div>
          <p className="text-xs text-slate-500 mb-4">Un file .xlsx separato per ciascuna sezione, pronto da stampare o condividere.</p>
          <div className="flex flex-col gap-2">
            <Button variant="secondary" className="justify-start" onClick={exportPlayersExcel}>
              <Users className="w-4 h-4" /> Schede giocatori (.xlsx)
            </Button>
            <Button variant="secondary" className="justify-start" onClick={exportTrainingsExcel}>
              <Activity className="w-4 h-4" /> Storico presenze (.xlsx)
            </Button>
            <Button variant="secondary" className="justify-start" onClick={exportMatchesExcel}>
              <Trophy className="w-4 h-4" /> Riepilogo partite (.xlsx)
            </Button>
            <Button variant="secondary" className="justify-start" onClick={exportStatisticsExcel}>
              <TrendingUp className="w-4 h-4" /> Statistiche giocatori (.xlsx)
            </Button>
          </div>
        </Card>
      </div>

      <Card className="p-4 flex items-start gap-3">
        <Info className="w-4.5 h-4.5 text-sky-400 shrink-0 mt-0.5" />
        <p className="text-xs text-slate-400">
          I dati vengono salvati automaticamente ad ogni modifica. Ti consigliamo comunque di scaricare un backup JSON periodicamente,
          soprattutto prima di importare dati o cambiare modalità di condivisione.
        </p>
      </Card>
    </div>
  );
}

function buildPitchHTML(formation, assignments, players) {
  const markers = formation.positions
    .map((pos) => {
      const player = players.find((p) => p.id === assignments[pos.id]);
      const label = player ? `${player.number ?? "-"}` : pos.label;
      const name = player ? shortName(player.name) : "Vuoto";
      return `<div class="marker" style="left:${pos.x}%; top:${pos.y}%;"><div class="dot">${label}</div>${name}</div>`;
    })
    .join("");
  return `<div class="pitch">${markers}</div>`;
}

// Costruisce un foglio Excel con layout professionale: titolo, sottotitolo,
// intestazioni raggruppate per area (con merge) e riquadro fisso sull'intestazione.
// Nota: la libreria SheetJS gratuita usata nell'app non supporta colori/font delle celle
// (funzionalità riservata alla versione Pro) — qui miglioriamo struttura, raggruppamenti,
// larghezze colonne e blocco delle intestazioni, che restano invece pienamente supportati.
function buildProfessionalSheet({ title, subtitle, groups, headers, rows, colWidths }) {
  const totalCols = headers.length;
  const aoa = [[title], [subtitle || ""]];
  const groupRow = new Array(totalCols).fill("");
  const merges = [];
  let colIdx = 0;
  (groups || []).forEach((g) => {
    groupRow[colIdx] = g.label;
    if (g.span > 1) merges.push({ s: { r: 2, c: colIdx }, e: { r: 2, c: colIdx + g.span - 1 } });
    colIdx += g.span;
  });
  aoa.push(groupRow);
  aoa.push(headers);
  rows.forEach((r) => aoa.push(r));

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: Math.max(0, totalCols - 1) } });
  merges.push({ s: { r: 1, c: 0 }, e: { r: 1, c: Math.max(0, totalCols - 1) } });
  ws["!merges"] = merges;
  ws["!cols"] = colWidths || headers.map(() => ({ wch: 14 }));
  ws["!rows"] = [{ hpt: 22 }, { hpt: 16 }, { hpt: 18 }, { hpt: 18 }];
  ws["!freeze"] = { xSplit: 0, ySplit: 4, topLeftCell: XLSX.utils.encode_cell({ r: 4, c: 0 }), activePane: "bottomLeft", state: "frozen" };
  return ws;
}

// Formatta la data corrente come es. "12ago26" (gg + mese abbreviato IT + aa),
// usata per marcare i file di backup con la data di esportazione.
function formatBackupDateSuffix(date = new Date()) {
  const months = ["gen", "feb", "mar", "apr", "mag", "giu", "lug", "ago", "set", "ott", "nov", "dic"];
  const day = String(date.getDate()).padStart(2, "0");
  const month = months[date.getMonth()];
  const year = String(date.getFullYear()).slice(-2);
  return `${day}${month}${year}`;
}

function downloadBlob(content, filename, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Genera e scarica la scheda di un singolo esercizio: immagine (se presente) in cima,
// poi tutte le informazioni, vincolata a restare su una sola pagina in stampa.
function downloadExerciseSheet(ex) {
  let body = "";
  if (ex.image) {
    // Immagine a "piena pagina": esce dai margini laterali del contenuto (16px)
    // e non ha un limite di altezza, così il testo disegnato dentro l'immagine
    // (schema, legenda, ecc.) resta leggibile invece di essere rimpicciolito.
    body += `<div style="page-break-inside: avoid;"><img src="${ex.image}" alt="${ex.title}" style="display:block; width: calc(100% + 32px); margin: 0 -16px 12px -16px; object-fit: contain;" /></div>`;
  }
  body += `<h1 style="margin-top:0;">${ex.title || "Esercizio"}</h1>`;
  body += `<p><span class="badge">${ex.type || "Tecnica"}</span> <strong>Tempo:</strong> ${ex.time || "--"}</p>`;
  if (ex.goal) body += `<h2>Obiettivo</h2><p>${ex.goal}</p>`;
  if (ex.description) body += `<h2>Descrizione</h2><p>${ex.description}</p>`;
  downloadPrintableHTML(`esercizio-${ex.title || "senza-titolo"}.html`, `Esercizio — ${ex.title || ""}`, `<div style="page-break-after: avoid; page-break-inside: avoid;">${body}</div>`, 1100);
}

// Converte una dataURL (base64, come sono salvate le immagini e i documenti
// nell'app) in un vero oggetto File: la Web Share API richiede File reali,
// non stringhe, per poter condividere allegati.
function dataUrlToFile(dataUrl, filename) {
  const [header, base64] = dataUrl.split(",");
  const mimeMatch = header.match(/data:(.*?);base64/);
  const mime = mimeMatch ? mimeMatch[1] : "application/octet-stream";
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return new File([bytes], filename, { type: mime });
}

// Prova a condividere uno o più file tramite il selettore nativo del dispositivo
// (Web Share API). Su smartphone/tablet questo mostra WhatsApp tra le app
// disponibili: selezionandolo, è WhatsApp stesso a far scegliere il gruppo o
// contatto di destinazione — è la via più diretta oggi possibile, perché
// WhatsApp non offre un modo per un sito web di inviare file a un gruppo
// specifico senza che sia l'utente a sceglierlo.
// Ritorna true se la condivisione è partita (o annullata volontariamente),
// false se questo dispositivo/browser non la supporta (tipico su desktop),
// nel qual caso va usato un ripiego (scaricare il file + link WhatsApp con testo).
async function shareFilesNative(files, meta = {}) {
  try {
    if (typeof navigator !== "undefined" && navigator.canShare && navigator.canShare({ files })) {
      await navigator.share({ files, title: meta.title, text: meta.text });
      return true;
    }
  } catch (e) {
    if (e?.name === "AbortError") return true; // l'utente ha annullato: non è un errore da segnalare
  }
  return false;
}

// Ripiego quando la condivisione diretta del file non è supportata: apre
// WhatsApp (app o Web) con un messaggio pronto, così l'utente sceglie il
// gruppo e allega a mano il file che nel frattempo è stato scaricato.
function openWhatsAppFallback(text) {
  window.open(`https://wa.me/?text=${encodeURIComponent(text)}`, "_blank");
}

// Funzione condivisa da tutti i pulsanti "Condividi": tenta la condivisione
// nativa del file (WhatsApp compreso) e, se non disponibile, scarica il file
// e apre comunque WhatsApp con un messaggio pronto per l'invio manuale.
async function shareOrFallback({ dataUrl, filename, label, showToast }) {
  const file = dataUrlToFile(dataUrl, filename);
  const shared = await shareFilesNative([file], { title: label, text: label });
  if (shared) return;
  const a = document.createElement("a");
  a.href = dataUrl;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  openWhatsAppFallback(`${label} — allega il file appena scaricato`);
  showToast?.("Condivisione diretta non disponibile su questo browser: file scaricato, allegalo su WhatsApp manualmente.", "error");
}

// Genera e scarica la scheda completa di un Focus Tecnico: tutte le immagini
// degli esercizi in sequenza, ciascuna a piena pagina, seguita dai dettagli.
function downloadFocusSheet(ft) {
  let body = `<h1 style="margin-top:0;">${ft.title || "Focus Tecnico"}</h1>`;
  body += `<p><strong>Durata totale:</strong> ${totalFocusMinutes(ft)} min</p>`;
  (ft.exercises || []).forEach((ex, i) => {
    body += `<div style="page-break-inside: avoid; margin-top: 24px;">`;
    if (ex.image) {
      body += `<img src="${ex.image}" alt="${ex.title || ""}" style="display:block; width: calc(100% + 32px); margin: 0 -16px 10px -16px; object-fit: contain;" />`;
    }
    body += `<h2>${i + 1}. ${ex.title || "Esercizio"}</h2>`;
    body += `<p><span class="badge">${ex.type || "Tecnica"}</span> <strong>Tempo:</strong> ${ex.time || "--"}</p>`;
    if (ex.goal) body += `<p><strong>Obiettivo:</strong> ${ex.goal}</p>`;
    if (ex.description) body += `<p>${ex.description}</p>`;
    body += `</div>`;
  });
  downloadPrintableHTML(`focus-${ft.title || "senza-titolo"}.html`, `Focus Tecnico — ${ft.title || ""}`, body, 1100);
}

// Blocco HTML condiviso per un singolo esercizio all'interno di una scheda
// multi-esercizio (Categoria o PlayBook completo): immagine a piena larghezza
// seguita dai dettagli, con interruzione di pagina prima di ogni esercizio
// tranne il primo, per avere un esercizio per pagina in stampa.
function exerciseBlockHtml(ex, indexLabel, isFirst) {
  let block = `<div style="page-break-inside: avoid; ${isFirst ? "" : "page-break-before: always;"} margin-top: ${isFirst ? "0" : "24px"};">`;
  if (ex.image) {
    block += `<img src="${ex.image}" alt="${ex.title || ""}" style="display:block; width: calc(100% + 32px); margin: 0 -16px 10px -16px; object-fit: contain;" />`;
  }
  block += `<h2>${indexLabel}. ${ex.title || "Esercizio"}</h2>`;
  block += `<p><span class="badge">${ex.type || "Tecnica"}</span> <strong>Tempo:</strong> ${ex.time || "--"}${ex.category ? ` · ${ex.category}` : ""}</p>`;
  if (ex.goal) block += `<p><strong>Obiettivo:</strong> ${ex.goal}</p>`;
  if (ex.description) block += `<p>${ex.description}</p>`;
  block += `</div>`;
  return block;
}

// Scarica tutti gli esercizi di una singola tipologia (es. tutti quelli "Tecnica").
function downloadCategorySheet(typeLabel, items) {
  let body = `<h1 style="margin-top:0;">Categoria: ${typeLabel}</h1><p>${items.length} esercizi</p>`;
  items.forEach((ex, i) => {
    body += exerciseBlockHtml(ex, i + 1, i === 0);
  });
  downloadPrintableHTML(`categoria-${typeLabel}.html`, `Categoria — ${typeLabel}`, body, 1100);
}

// Scarica l'intero PlayBook: tutti gli esercizi presenti, raggruppati per
// tipologia nello stesso ordine mostrato nella schermata PlayBook.
function downloadPlayBookSheet(allExercises, typesOrder) {
  let body = `<h1 style="margin-top:0;">PlayBook completo</h1><p>${allExercises.length} esercizi totali</p>`;
  let counter = 0;
  let firstBlock = true;
  (typesOrder || []).forEach((type) => {
    const items = allExercises.filter((ex) => (ex.type || "ND") === type);
    if (items.length === 0) return;
    body += `<h1 style="${firstBlock ? "" : "page-break-before: always;"} margin-top: 32px;">${type}</h1>`;
    items.forEach((ex) => {
      counter += 1;
      body += exerciseBlockHtml(ex, counter, true);
    });
    firstBlock = false;
  });
  downloadPrintableHTML("playbook-completo.html", "PlayBook completo", body, 1100);
}

// Visualizzatore "esploso" di un Focus Tecnico: mostra un esercizio alla volta
// a piena schermata con immagine grande, con navigazione avanti/indietro e la
// possibilità di scaricare (o condividere) tutte le immagini della sessione.
function FocusDetailViewer({ focus, onClose, showToast }) {
  const exercisesWithContent = (focus.exercises || []);
  const [index, setIndex] = useState(0);
  const current = exercisesWithContent[index];
  if (!current) return null;

  function goTo(delta) {
    setIndex((i) => Math.max(0, Math.min(exercisesWithContent.length - 1, i + delta)));
  }

  function downloadAllImages() {
    const withImages = exercisesWithContent.filter((ex) => ex.image);
    if (withImages.length === 0) return showToast?.("Nessuna immagine presente in questo Focus", "error");
    const baseName = (focus.title || "focus").replace(/[^a-z0-9]+/gi, "-");
    const summaryUrl = generateFocusSummaryImage(focus);
    const a0 = document.createElement("a");
    a0.href = summaryUrl;
    a0.download = `${baseName}-0-riepilogo.png`;
    document.body.appendChild(a0);
    a0.click();
    document.body.removeChild(a0);
    withImages.forEach((ex, i) => {
      setTimeout(() => {
        const a = document.createElement("a");
        a.href = ex.image;
        a.download = `${baseName}-${i + 1}-${(ex.title || "esercizio").replace(/[^a-z0-9]+/gi, "-")}.jpg`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
      }, (i + 1) * 350); // piccolo intervallo tra un download e l'altro: alcuni browser bloccano download multipli simultanei
    });
    showToast?.(`Download di ${withImages.length + 1} file avviato (riepilogo + immagini)`);
  }

  async function shareAllImages() {
    const withImages = exercisesWithContent.filter((ex) => ex.image);
    if (withImages.length === 0) return showToast?.("Nessuna immagine presente in questo Focus", "error");
    const summaryFile = dataUrlToFile(generateFocusSummaryImage(focus), "0-riepilogo-focus.png");
    const exerciseFiles = withImages.map((ex, i) =>
      dataUrlToFile(ex.image, `${(ex.title || `esercizio-${i + 1}`).replace(/[^a-z0-9]+/gi, "-")}.jpg`)
    );
    const files = [summaryFile, ...exerciseFiles];
    const shared = await shareFilesNative(files, { title: focus.title, text: focus.title });
    if (!shared) {
      downloadAllImages();
      openWhatsAppFallback(`${focus.title || "Focus Tecnico"} — allega le immagini appena scaricate`);
      showToast?.("Condivisione diretta non disponibile: immagini scaricate, allegale su WhatsApp manualmente.", "error");
    }
  }

  return (
    <div className="fixed inset-0 bg-black/95 flex flex-col" style={{ zIndex: 100 }}>
      <div className="flex items-center justify-between gap-2 px-4 py-3 border-b border-white/10 flex-wrap">
        <div className="min-w-0">
          <p className="text-sm font-bold text-white truncate">{focus.title || "Focus Tecnico"}</p>
          <p className="text-[11px] text-slate-400">Esercizio {index + 1} di {exercisesWithContent.length}</p>
        </div>
        <div className="flex items-center gap-1.5 shrink-0">
          <button onClick={shareAllImages} className="flex items-center gap-1.5 rounded-lg px-2.5 py-1.5 text-xs font-medium bg-white/10 hover:bg-white/20 text-white" title="Condividi tutte le immagini">
            <Share2 className="w-3.5 h-3.5" /> <span className="hidden sm:inline">Condividi tutte</span>
          </button>
          <button onClick={downloadAllImages} className="flex items-center gap-1.5 rounded-lg px-2.5 py-1.5 text-xs font-medium bg-white/10 hover:bg-white/20 text-white" title="Scarica tutte le immagini">
            <Download className="w-3.5 h-3.5" /> <span className="hidden sm:inline">Scarica tutte</span>
          </button>
          <button onClick={onClose} className="rounded-lg p-2 hover:bg-white/10 text-white">
            <X className="w-5 h-5" />
          </button>
        </div>
      </div>

      <div className="flex-1 flex items-center justify-center relative px-4 min-h-0">
        {index > 0 && (
          <button onClick={() => goTo(-1)} className="absolute left-2 sm:left-4 rounded-full p-2 bg-white/10 hover:bg-white/20 text-white">
            <ChevronLeft className="w-6 h-6" />
          </button>
        )}
        <div className="max-w-3xl w-full flex flex-col items-center gap-3 py-4 overflow-y-auto" style={{ maxHeight: "100%" }}>
          {current.image ? (
            <img src={current.image} alt={current.title} className="max-w-full object-contain rounded-lg" style={{ maxHeight: "52vh" }} />
          ) : (
            <div className="w-full aspect-video rounded-lg bg-white/5 flex items-center justify-center text-slate-500 text-sm">
              Nessuna immagine per questo esercizio
            </div>
          )}
          <div className="w-full text-left px-1">
            <div className="flex items-center gap-2 flex-wrap mb-1">
              {current.type && <Badge className={EXERCISE_TYPE_STYLES[current.type] || EXERCISE_TYPE_STYLES.Tecnica}>{current.type}</Badge>}
              <p className="text-base font-bold text-white">{current.title || "Esercizio"}</p>
              <span className="text-xs text-slate-400">· {current.time || "--"}</span>
            </div>
            {current.goal && <p className="text-sm text-emerald-400 italic mb-1">Obiettivo: {current.goal}</p>}
            {current.description && <p className="text-sm text-slate-300">{current.description}</p>}
          </div>
        </div>
        {index < exercisesWithContent.length - 1 && (
          <button onClick={() => goTo(1)} className="absolute right-2 sm:right-4 rounded-full p-2 bg-white/10 hover:bg-white/20 text-white">
            <ChevronRight className="w-6 h-6" />
          </button>
        )}
      </div>

      <div className="flex items-center justify-center gap-1.5 py-3 border-t border-white/10 overflow-x-auto px-4">
        {exercisesWithContent.map((ex, i) => (
          <button
            key={ex.id || i}
            onClick={() => setIndex(i)}
            className={`w-2 h-2 rounded-full shrink-0 ${i === index ? "bg-emerald-400" : "bg-white/20"}`}
            title={ex.title}
          />
        ))}
      </div>
    </div>
  );
}

// Anteprima di un documento del Dossier senza doverlo scaricare: immagini e
// PDF vengono mostrati direttamente; per gli altri formati (Word, Excel, ecc.)
// il browser non può renderizzare in linea, quindi si propone il download.
function DossierPreviewModal({ doc, onClose, showToast }) {
  if (!doc) return null;
  const isImage = doc.fileType?.startsWith("image/");
  const isPdf = doc.fileType === "application/pdf";

  function download() {
    const a = document.createElement("a");
    a.href = doc.dataUrl;
    a.download = doc.fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  return (
    <div className="fixed inset-0 bg-black/90 z-[100] flex flex-col">
      <div className="flex items-center justify-between gap-2 px-4 py-3 border-b border-white/10">
        <div className="min-w-0">
          <p className="text-sm font-bold text-white truncate">{doc.title}</p>
          <p className="text-[11px] text-slate-400 truncate">{doc.fileName}</p>
        </div>
        <div className="flex items-center gap-1.5 shrink-0">
          <button
            onClick={() => shareOrFallback({ dataUrl: doc.dataUrl, filename: doc.fileName, label: doc.title, showToast })}
            className="flex items-center gap-1.5 rounded-lg px-2.5 py-1.5 text-xs font-medium bg-white/10 hover:bg-white/20 text-white"
          >
            <Share2 className="w-3.5 h-3.5" /> <span className="hidden sm:inline">Condividi</span>
          </button>
          <button onClick={download} className="flex items-center gap-1.5 rounded-lg px-2.5 py-1.5 text-xs font-medium bg-white/10 hover:bg-white/20 text-white">
            <Download className="w-3.5 h-3.5" /> <span className="hidden sm:inline">Scarica</span>
          </button>
          <button onClick={onClose} className="rounded-lg p-2 hover:bg-white/10 text-white">
            <X className="w-5 h-5" />
          </button>
        </div>
      </div>
      <div className="flex-1 flex items-center justify-center overflow-auto p-4">
        {isImage ? (
          <img src={doc.dataUrl} alt={doc.title} className="max-w-full max-h-full object-contain rounded-lg" />
        ) : isPdf ? (
          <iframe src={doc.dataUrl} title={doc.title} className="w-full h-full rounded-lg bg-white" style={{ minHeight: "70vh" }} />
        ) : (
          <div className="text-center text-slate-400 text-sm max-w-sm">
            <FileText className="w-10 h-10 mx-auto mb-3 text-slate-600" />
            Anteprima non disponibile per questo tipo di file ({doc.fileType || "sconosciuto"}).
            <br />
            Usa "Scarica" per aprirlo con l'app giusta sul tuo dispositivo.
          </div>
        )}
      </div>
    </div>
  );
}

// Genera un'immagine di riepilogo del Focus Tecnico (titolo, durata totale ed
// elenco esercizi con tipo/tempo/obiettivo) da anteporre alle immagini quando
// si condivide o si scarica l'intera sessione: chi riceve vede subito di cosa
// si tratta prima delle singole foto degli esercizi.
function generateFocusSummaryImage(focus) {
  const exercisesList = focus.exercises || [];
  const width = 900;
  const padding = 44;
  const headerHeight = 120;
  const rowHeights = exercisesList.map((ex) => (ex.goal ? 74 : 50));
  const height = headerHeight + rowHeights.reduce((a, b) => a + b, 0) + padding * 2;

  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = Math.max(height, 260);
  const ctx = canvas.getContext("2d");

  ctx.fillStyle = "#0f172a";
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  ctx.fillStyle = "#f1f5f9";
  ctx.font = "bold 34px Arial, sans-serif";
  ctx.fillText(focus.title || "Focus Tecnico", padding, 58);

  ctx.fillStyle = "#34d399";
  ctx.font = "600 19px Arial, sans-serif";
  ctx.fillText(`Durata totale: ${totalFocusMinutes(focus)} min · ${exercisesList.length} esercizi`, padding, 88);

  ctx.strokeStyle = "rgba(255,255,255,0.15)";
  ctx.beginPath();
  ctx.moveTo(padding, 104);
  ctx.lineTo(width - padding, 104);
  ctx.stroke();

  let y = headerHeight + padding - 30;
  exercisesList.forEach((ex, i) => {
    ctx.fillStyle = "#34d399";
    ctx.font = "bold 21px Arial, sans-serif";
    ctx.fillText(`${i + 1}. ${ex.title || "Esercizio"}`, padding, y);

    ctx.fillStyle = "#94a3b8";
    ctx.font = "16px Arial, sans-serif";
    ctx.fillText(`${ex.type || "Tecnica"} · ${ex.time || "--"}`, padding, y + 25);

    if (ex.goal) {
      ctx.fillStyle = "#cbd5e1";
      ctx.font = "italic 15px Arial, sans-serif";
      ctx.fillText(`Obiettivo: ${ex.goal}`, padding, y + 48);
    }

    y += rowHeights[i];
  });

  return canvas.toDataURL("image/png");
}

// Genera un'immagine di convocazione (sfondo verde chiaro, stile "locandina")
// pronta da condividere su WhatsApp: intestazione con avversario/data/luogo e
// elenco dei convocati con solo il nome (senza numero di maglia né ruolo, su
// richiesta, per un messaggio più semplice da leggere per i genitori).
// Riduce la dimensione del font finché il testo non entra nella larghezza
// massima disponibile, per evitare che nomi squadra lunghi escano dal bordo.
function fitBoldFontSize(ctx, text, maxWidth, startSize, minSize = 16) {
  let size = startSize;
  while (size > minSize) {
    ctx.font = `bold ${size}px Arial, sans-serif`;
    if (ctx.measureText(text).width <= maxWidth) break;
    size -= 2;
  }
  return size;
}

function generateConvocationImage(match, players, clubLabel) {
  const convocati = players.filter((p) => (match.convocati || []).includes(p.id));
  const width = 800;
  const padding = 48;
  const headerHeight = 210;
  const rowHeight = 42;
  const height = headerHeight + Math.max(convocati.length, 1) * rowHeight + padding;

  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");
  const maxLineWidth = width - padding * 2;

  // Sfondo verde chiaro, come richiesto
  ctx.fillStyle = "#dcfce7";
  ctx.fillRect(0, 0, width, height);

  // Tipo partita (Campionato/Amichevole/Torneo), in evidenza in alto
  ctx.fillStyle = "#15803d";
  ctx.font = "700 15px Arial, sans-serif";
  const typeLabel = [match.matchType, match.matchType === "Torneo" ? match.tournamentName : null].filter(Boolean).join(" · ");
  ctx.fillText((typeLabel || "Partita").toUpperCase(), padding, 34);

  // Squadre: stessa dimensione di font per entrambi i nomi
  const teamsLine = `${clubLabel || "Squadra"}  vs  ${match.opponent || ""}`;
  const teamsFontSize = fitBoldFontSize(ctx, teamsLine, maxLineWidth, 32);
  ctx.fillStyle = "#052e1a";
  ctx.font = `bold ${teamsFontSize}px Arial, sans-serif`;
  ctx.fillText(teamsLine, padding, 82);

  ctx.fillStyle = "#166534";
  ctx.font = "500 18px Arial, sans-serif";
  const infoLine = [formatDate(match.date), match.time, match.venue].filter(Boolean).join(" · ");
  ctx.fillText(infoLine, padding, 112);

  ctx.strokeStyle = "rgba(5,46,26,0.25)";
  ctx.beginPath();
  ctx.moveTo(padding, 134);
  ctx.lineTo(width - padding, 134);
  ctx.stroke();

  ctx.fillStyle = "#052e1a";
  ctx.font = "bold 20px Arial, sans-serif";
  ctx.fillText(`Convocati (${convocati.length})`, padding, 166);

  let y = headerHeight + 20;
  if (convocati.length === 0) {
    ctx.fillStyle = "#166534";
    ctx.font = "italic 16px Arial, sans-serif";
    ctx.fillText("Nessun convocato ancora selezionato", padding, y);
  } else {
    convocati.forEach((p, i) => {
      ctx.fillStyle = "#052e1a";
      ctx.font = "600 20px Arial, sans-serif";
      ctx.fillText(`${i + 1}. ${p.name}`, padding, y);
      y += rowHeight;
    });
  }

  return canvas.toDataURL("image/png");
}

// Overlay per ingrandire un'immagine con pulsante di download
function ImageLightbox({ src, alt, onClose }) {
  if (!src) return null;
  function handleDownload() {
    const a = document.createElement("a");
    a.href = src;
    a.download = (alt || "immagine").replace(/[^a-z0-9]+/gi, "-") + ".jpg";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }
  return (
    <div
      className="fixed inset-0 bg-black/90 flex flex-col items-center justify-center p-4"
      style={{ zIndex: 100 }}
      onClick={onClose}
    >
      <button onClick={onClose} className="absolute top-4 right-4 text-white rounded-lg p-2 hover:bg-white/10">
        <X className="w-6 h-6" />
      </button>
      <img
        src={src}
        alt={alt}
        className="max-w-full object-contain rounded-lg"
        style={{ maxHeight: "75vh" }}
        onClick={(e) => e.stopPropagation()}
      />
      <button
        onClick={(e) => {
          e.stopPropagation();
          handleDownload();
        }}
        className="mt-4 flex items-center gap-2 rounded-xl bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-semibold px-4 py-2.5 text-sm"
      >
        <Download className="w-4 h-4" /> Scarica immagine
      </button>
    </div>
  );
}

// Genera un file HTML autonomo e stampabile (funziona anche quando window.print()
// è bloccato dalla sandbox dell'artifact): l'utente lo apre nel browser e da lì
// può stampare o salvare come PDF.
function downloadPrintableHTML(filename, docTitle, bodyHtml, maxWidth = 800) {
  const html = `<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="UTF-8" />
<title>${docTitle}</title>
<style>
  body { font-family: -apple-system, Arial, sans-serif; color: #111; max-width: ${maxWidth}px; margin: 24px auto; padding: 0 16px; }
  h1 { font-size: 20px; margin-bottom: 4px; }
  h2 { font-size: 16px; margin: 18px 0 6px; }
  p { font-size: 13px; line-height: 1.5; margin: 2px 0; }
  table { border-collapse: collapse; width: 100%; margin-top: 8px; }
  td, th { border: 1px solid #ccc; padding: 6px 8px; font-size: 12px; text-align: left; }
  .pitch { position: relative; width: 320px; height: 460px; background: #166534; border-radius: 12px; margin: 12px 0; border: 2px solid #14532d; }
  .pitch .line { position: absolute; left: 2%; right: 2%; top: 50%; border-top: 1px solid rgba(255,255,255,0.5); }
  .pitch .marker { position: absolute; transform: translate(-50%, -50%); text-align: center; font-size: 10px; color: #fff; width: 70px; }
  .pitch .dot { width: 30px; height: 30px; border-radius: 50%; background: #10b981; border: 2px solid #fff; margin: 0 auto 2px; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 11px; color: #052e1a; }
  .page-break { page-break-after: always; }
  .badge { display: inline-block; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 600; background: #e2e8f0; color: #334155; margin-right: 4px; }
  ul, ol { font-size: 13px; padding-left: 20px; }
  li { margin-bottom: 6px; }
  .print-btn { margin: 12px 0; padding: 8px 16px; background: #10b981; color: #052e1a; border: none; border-radius: 8px; font-weight: 600; cursor: pointer; }
  @media print { .print-btn { display: none; } }
</style>
</head>
<body>
<button class="print-btn" onclick="window.print()">Stampa / Salva come PDF</button>
${bodyHtml}
</body>
</html>`;
  downloadBlob(html, filename, "text/html");
}

/* ============================================================
   SEZIONE IMPOSTAZIONI
   ============================================================ */

function SettingsSection({ season, updateSeason, sharedMode, onToggleShared, seasons, setSeasons, activeSeasonId, setActiveSeasonId, showToast }) {
  const [form, setForm] = useState({
    name: season.name,
    teamName: season.teamName,
    leva: season.leva,
    teamFormat: season.teamFormat || "",
    colorPrimary: season.colorPrimary || "#10b981",
    colorSecondary: season.colorSecondary || "#0f172a",
    logoUrl: season.logoUrl || null,
  });
  const [confirmReset, setConfirmReset] = useState(false);
  const [confirmDeleteSeason, setConfirmDeleteSeason] = useState(false);
  const [logoError, setLogoError] = useState("");
  const logoInputRef = React.useRef(null);

  useEffect(() => {
    setForm({
      name: season.name,
      teamName: season.teamName,
      leva: season.leva,
      teamFormat: season.teamFormat || "",
      colorPrimary: season.colorPrimary || "#10b981",
      colorSecondary: season.colorSecondary || "#0f172a",
      logoUrl: season.logoUrl || null,
    });
  }, [season.id]);

  async function handleLogoChange(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setLogoError("");
    try {
      const dataUrl = await resizeImageFile(file, 400, 0.9);
      setForm((f) => ({ ...f, logoUrl: dataUrl }));
    } catch (err) {
      setLogoError("Immagine non valida, riprova con un altro file.");
    }
    e.target.value = "";
  }

  function save() {
    updateSeason(() => ({ ...form }));
    showToast("Impostazioni salvate");
  }

  function resetSeasonData() {
    updateSeason(() => ({
      players: [],
      trainings: [],
      matches: [],
      championship: { fase1: null, fase2: null, faseFinale: null },
    }));
    setConfirmReset(false);
    showToast("Dati della stagione azzerati (Giocatori, Allenamenti, Partite, Campionato)");
  }

  function deleteSeason() {
    const remaining = seasons.filter((s) => s.id !== activeSeasonId);
    if (remaining.length === 0) {
      showToast("Non puoi eliminare l'unica stagione esistente", "error");
      setConfirmDeleteSeason(false);
      return;
    }
    setSeasons(remaining);
    setActiveSeasonId(remaining[0].id);
    setConfirmDeleteSeason(false);
    showToast("Stagione eliminata");
  }

  return (
    <div className="max-w-2xl">
      <SectionTitle eyebrow="Configurazione" title="Impostazioni" icon={Settings} />

      <Card className="p-5 mb-5">
        <p className="text-sm font-bold text-slate-100 mb-4">Struttura: Stagione · Squadra · Leva</p>
        <Field label="Nome stagione">
          <input className={inputClass} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="Es. Stagione 2026-27" />
        </Field>
        <Field label="Nome squadra">
          <input className={inputClass} value={form.teamName} onChange={(e) => setForm({ ...form, teamName: e.target.value })} placeholder="Es. Prima Squadra" />
        </Field>
        <Field label="Logo / Stemma squadra">
          <div className="flex items-center gap-3">
            <div className="w-14 h-14 rounded-xl overflow-hidden border border-white/10 bg-slate-900 flex items-center justify-center shrink-0">
              {form.logoUrl ? (
                <img src={form.logoUrl} alt="Logo squadra" className="w-full h-full object-cover" />
              ) : (
                <Shield className="w-6 h-6 text-slate-600" />
              )}
            </div>
            <div className="flex gap-2">
              <Button type="button" variant="secondary" onClick={() => logoInputRef.current?.click()}>
                {form.logoUrl ? "Sostituisci" : "Carica logo"}
              </Button>
              {form.logoUrl && (
                <Button type="button" variant="ghost" onClick={() => setForm((f) => ({ ...f, logoUrl: null }))}>
                  Rimuovi
                </Button>
              )}
            </div>
            <input ref={logoInputRef} type="file" accept="image/*" className="hidden" onChange={handleLogoChange} />
          </div>
          {logoError && <p className="text-[11px] text-rose-400 mt-1">{logoError}</p>}
          <p className="text-[11px] text-slate-500 mt-1">Comparirà in alto, accanto al nome squadra e stagione.</p>
        </Field>
        <Field label="Leva calcistica">
          <input className={inputClass} value={form.leva} onChange={(e) => setForm({ ...form, leva: e.target.value })} placeholder="Es. Leva 2012 / Giovanissimi" />
        </Field>
        <Field label="Tipo campionato">
          <select className={inputClass} value={form.teamFormat} onChange={(e) => setForm({ ...form, teamFormat: e.target.value })}>
            <option value="">— Non specificato —</option>
            {TEAM_FORMAT_OPTIONS.map((f) => (
              <option key={f} value={f}>{f}</option>
            ))}
          </select>
        </Field>
        <Field label="Colore sociale primario">
          <ColorSwatchPicker value={form.colorPrimary} onChange={(hex) => setForm({ ...form, colorPrimary: hex })} />
        </Field>
        <Field label="Colore sociale secondario">
          <ColorSwatchPicker value={form.colorSecondary} onChange={(hex) => setForm({ ...form, colorSecondary: hex })} />
        </Field>
        <Button onClick={save}><Save className="w-4 h-4" /> Salva Impostazioni</Button>
      </Card>

      <Card className="p-5 mb-5">
        <p className="text-sm font-bold text-slate-100 mb-2">Condivisione dati con lo staff</p>
        <p className="text-xs text-slate-500 mb-4">
          In modalità condivisa, chiunque apra questo stesso artifact vede e modifica gli stessi dati: nessun login separato per collaboratore.
          Utile per uno staff ristretto con cui condividi il link della chat.
        </p>
        <div className="flex items-center justify-between rounded-xl border border-white/10 p-3.5">
          <div className="flex items-center gap-2.5">
            {sharedMode ? <Cloud className="w-4.5 h-4.5 text-emerald-400" /> : <CloudOff className="w-4.5 h-4.5 text-slate-500" />}
            <div>
              <p className="text-sm font-medium text-slate-200">{sharedMode ? "Modalità condivisa attiva" : "Modalità personale attiva"}</p>
              <p className="text-[11px] text-slate-500">{sharedMode ? "Dati visibili a chi ha il link" : "Dati visibili solo a te"}</p>
            </div>
          </div>
          <button
            onClick={onToggleShared}
            className={`w-12 h-7 rounded-full transition-colors relative ${sharedMode ? "bg-emerald-500" : "bg-slate-700"}`}
          >
            <span className={`absolute top-1 w-5 h-5 rounded-full bg-white transition-transform ${sharedMode ? "translate-x-6" : "translate-x-1"}`} />
          </button>
        </div>
      </Card>

      <Card className="p-5 border-rose-500/20">
        <p className="text-sm font-bold text-rose-400 mb-4">Zona pericolosa</p>

        <div className="flex items-center justify-between mb-4 gap-3">
          <div>
            <p className="text-sm text-slate-200">Azzera dati di questa stagione</p>
            <p className="text-[11px] text-slate-500">Rimuove giocatori, allenamenti e partite di "{season.name}".</p>
          </div>
          {confirmReset ? (
            <div className="flex gap-2 shrink-0">
              <Button variant="secondary" onClick={() => setConfirmReset(false)}>Annulla</Button>
              <Button variant="danger" onClick={resetSeasonData}>Conferma</Button>
            </div>
          ) : (
            <Button variant="danger" onClick={() => setConfirmReset(true)}><Trash2 className="w-4 h-4" /> Azzera</Button>
          )}
        </div>

        <div className="flex items-center justify-between gap-3">
          <div>
            <p className="text-sm text-slate-200">Elimina stagione "{season.name}"</p>
            <p className="text-[11px] text-slate-500">Rimuove definitivamente la stagione e tutti i suoi dati.</p>
          </div>
          {confirmDeleteSeason ? (
            <div className="flex gap-2 shrink-0">
              <Button variant="secondary" onClick={() => setConfirmDeleteSeason(false)}>Annulla</Button>
              <Button variant="danger" onClick={deleteSeason}>Conferma</Button>
            </div>
          ) : (
            <Button variant="danger" onClick={() => setConfirmDeleteSeason(true)}><Trash2 className="w-4 h-4" /> Elimina</Button>
          )}
        </div>
      </Card>
    </div>
  );
}
