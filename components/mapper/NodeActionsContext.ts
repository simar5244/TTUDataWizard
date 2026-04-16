"use client";

import { createContext, useContext } from "react";

export const NodeActionsContext = createContext<{ deleteNode: (id: string) => void }>({
  deleteNode: () => {},
});

export function useNodeActions() {
  return useContext(NodeActionsContext);
}
