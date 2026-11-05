<!-- The decision guide drawing, included by data-security.qmd and the deck. Its words are in data-security.qmd's
     .decision-words list too, for phones; change both together. -->

```{=html}
<svg class="decision-svg" viewBox="0 0 760 540" width="100%" style="max-width:760px" role="img" aria-labelledby="dg-title dg-desc" font-family="Inter, system-ui, -apple-system, 'Segoe UI', Arial, sans-serif">
  <title id="dg-title">Can I upload this data to Gemini?</title>
  <desc id="dg-desc">Four questions in order. One: signed in with a personal Google account? If yes, switch to your NC State account; the free public Gemini isn't approved for university data. Two: is any of it purple, such as Social Security numbers, passwords, keys or biometrics? If yes, don't put it in any AI tool. Three: could an agreement limit sharing, such as an IRB protocol, a data use agreement, a partner NDA, sponsor terms or export control? If yes, describe it only, with headers and made-up rows, and ask whoever signed the agreement first. Otherwise, for green, yellow or red data that nothing limits, you may upload it or describe it; if you upload, include only the columns you need.</desc>
  <defs>
    <marker id="dg-arrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="8" markerHeight="8" orient="auto-start-reverse">
      <path d="M0,0 L10,5 L0,10 z" fill="#57606A"/>
    </marker>
  </defs>
  <g font-size="17" fill="#1F2328">
    <!-- 1. account -->
    <rect x="20" y="20" width="330" height="92" rx="10" fill="#FFFFFF" stroke="#1F2328" stroke-width="2"/>
    <text x="36" y="60">Signed in with a personal</text><text x="36" y="82">Google account?</text>
    <line x1="352" y1="66" x2="404" y2="66" stroke="#57606A" stroke-width="2" marker-end="url(#dg-arrow)"/>
    <text x="378" y="56" font-size="15" font-weight="700" fill="#57606A" text-anchor="middle">Yes</text>
    <rect x="410" y="20" width="330" height="92" rx="10" fill="#FDEDE4" stroke="#D14905" stroke-width="2"/>
    <text x="426" y="49" font-weight="700">Switch to your NC State account.</text><text x="426" y="71">The free public Gemini isn't</text><text x="426" y="93">approved for university data.</text>
    <line x1="185" y1="114" x2="185" y2="156" stroke="#57606A" stroke-width="2" marker-end="url(#dg-arrow)"/>
    <text x="196" y="140" font-size="15" font-weight="700" fill="#57606A">No</text>
    <!-- 2. purple -->
    <rect x="20" y="162" width="330" height="72" rx="10" fill="#FFFFFF" stroke="#1F2328" stroke-width="2"/>
    <text x="36" y="192">Is any of it purple? SSNs,</text><text x="36" y="214">passwords, keys, biometrics</text>
    <line x1="352" y1="198" x2="404" y2="198" stroke="#57606A" stroke-width="2" marker-end="url(#dg-arrow)"/>
    <text x="378" y="188" font-size="15" font-weight="700" fill="#57606A" text-anchor="middle">Yes</text>
    <rect x="410" y="162" width="330" height="72" rx="10" fill="#FBE3E0" stroke="#CC0000" stroke-width="2"/>
    <text x="426" y="203" font-weight="700">Don't put it in any AI tool.</text>
    <line x1="185" y1="236" x2="185" y2="278" stroke="#57606A" stroke-width="2" marker-end="url(#dg-arrow)"/>
    <text x="196" y="262" font-size="15" font-weight="700" fill="#57606A">No</text>
    <!-- 3. agreements -->
    <rect x="20" y="284" width="330" height="116" rx="10" fill="#FFFFFF" stroke="#1F2328" stroke-width="2"/>
    <text x="36" y="314">Could an agreement limit sharing?</text><text x="36" y="336">IRB protocol, data use agreement,</text><text x="36" y="358">partner NDA, sponsor terms,</text><text x="36" y="380">export control</text>
    <line x1="352" y1="342" x2="404" y2="342" stroke="#57606A" stroke-width="2" marker-end="url(#dg-arrow)"/>
    <text x="378" y="332" font-size="15" font-weight="700" fill="#57606A" text-anchor="middle">Yes</text>
    <rect x="410" y="284" width="330" height="116" rx="10" fill="#E3E8F5" stroke="#4156A1" stroke-width="2"/>
    <text x="426" y="325" font-weight="700">Describe it only: headers and</text><text x="426" y="347">made-up rows. Ask whoever</text><text x="426" y="369">signed the agreement first.</text>
    <line x1="185" y1="402" x2="185" y2="444" stroke="#57606A" stroke-width="2" marker-end="url(#dg-arrow)"/>
    <text x="196" y="428" font-size="15" font-weight="700" fill="#57606A">No</text>
    <!-- 4. go ahead -->
    <rect x="20" y="450" width="330" height="72" rx="10" fill="#FFFFFF" stroke="#1F2328" stroke-width="2"/>
    <text x="36" y="480">Green, yellow or red data,</text><text x="36" y="502">and nothing limits sharing</text>
    <line x1="352" y1="486" x2="404" y2="486" stroke="#57606A" stroke-width="2" marker-end="url(#dg-arrow)"/>
    <rect x="410" y="450" width="330" height="72" rx="10" fill="#EEF1DC" stroke="#6F7D1C" stroke-width="2"/>
    <text x="426" y="480" font-weight="700">Upload it or describe it.</text><text x="426" y="502">Upload only the columns you need.</text>
  </g>
</svg>
```
