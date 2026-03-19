export type Principle = {
  title: string;
  details: string[];
};

export type BlogDraftRequest = {
  topic: string;
  audience: string;
  objective: string;
  tone: "Professional" | "Friendly" | "Confident" | "Educational";
  length: "Short" | "Medium" | "Long";
  cta: string;
  principles: Principle[];
};

export type BlogFinalRequest = {
  outlineText: string;
  topic: string;
  audience: string;
  objective: string;
  tone: "Professional" | "Friendly" | "Confident" | "Educational";
  cta: string;
  polishLevel: "Standard" | "Strong" | "Premium";
};
