export default function ResultRow({ label, value }) {
  return (
    <div className="flex justify-between text-sm py-1 border-b border-gray-100 last:border-none">
      <span className="text-gray-600">{label}</span>
      <span className="font-semibold text-gray-900">{value}</span>
    </div>
  );
}
