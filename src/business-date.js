const dateFormatter = new Intl.DateTimeFormat('en-CA', {
  timeZone: import.meta.env.VITE_BUSINESS_TIMEZONE || 'Asia/Kolkata',
  year: 'numeric', month: '2-digit', day: '2-digit'
});
export function businessDate(date = new Date()) {
  const parts = Object.fromEntries(dateFormatter.formatToParts(date).map(part => [part.type, part.value]));
  return `${parts.year}-${parts.month}-${parts.day}`;
}

