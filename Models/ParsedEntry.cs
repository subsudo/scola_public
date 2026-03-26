using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace VerlaufsakteApp.Models;

public class ParsedEntry : INotifyPropertyChanged
{
    private string _name = string.Empty;
    private string _status = string.Empty;
    private string _absenceReason = string.Empty;
    private string _remark = string.Empty;

    public string Name
    {
        get => _name;
        set => SetField(ref _name, value);
    }

    public string Status
    {
        get => _status;
        set => SetField(ref _status, value);
    }

    public string AbsenceReason
    {
        get => _absenceReason;
        set => SetField(ref _absenceReason, value);
    }

    public string Remark
    {
        get => _remark;
        set => SetField(ref _remark, value);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        return true;
    }
}
