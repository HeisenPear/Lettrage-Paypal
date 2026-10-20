type Child = Node | string | null | undefined | false;

/** Cree un element avec ses attributs et ses enfants (les enfants vides sont ignores). */
export function h<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  attrs: Record<string, string> = {},
  ...children: Child[]
): HTMLElementTagNameMap[K] {
  const el = document.createElement(tag);
  for (const [name, value] of Object.entries(attrs)) el.setAttribute(name, value);
  for (const child of children) if (child) el.append(child);
  return el;
}

/** Pastille d'indicateur : libelle en petit, valeur en gras. */
export function chip(label: string, value: string): HTMLElement {
  return h(
    "div",
    { class: "chip" },
    h("span", { class: "chip-label" }, label),
    h("span", { class: "chip-value" }, value)
  );
}

/** Liste deroulante a partir de couples [valeur, libelle]. */
export function selectOf(options: [string, string][], current: string): HTMLSelectElement {
  const select = h("select");
  for (const [value, label] of options) select.append(h("option", { value }, label));
  select.value = current;
  return select;
}
