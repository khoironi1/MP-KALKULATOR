export default function InputField({ label, name, value, onChange, suffix }) {
  return (
    <div className="flex flex-col gap-1">
      <label className="text-xs font-medium text-gray-700">{label}</label>
      <div className="flex items-center border rounded-lg overflow-hidden bg-white">
        <input
          type="number"
          name={name}
          value={value}
          onChange={onChange}
          className="w-full p-2 text-sm outline-none"
        />
        {suffix && <span className="px-2 text-sm bg-gray-100 text-gray-600 border-l">{suffix}</span>}
      </div>
    </div>
  );
}
