import { SpeedInsights } from "@vercel/speed-insights/next";

export const metadata = {
  title: "Virgin Voyages Dashboard",
  description: "Virgin Voyages Product, Equipment, Inventory and Schedule Dashboard",
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body>
        {children}
        <SpeedInsights />
      </body>
    </html>
  );
}
