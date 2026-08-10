using System.ComponentModel.DataAnnotations;

namespace FramesServer.Components.Pages.Home;

public class Step2Data
{
    [Required] [Range(1, 1000000)] public decimal? LastMonthAmt { get; set; }

    [Required] public Dictionary<DateOnly, decimal> Paid { get; set; } = [];
}
