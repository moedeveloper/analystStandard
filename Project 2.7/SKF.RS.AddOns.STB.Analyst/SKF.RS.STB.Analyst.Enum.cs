using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SKF.RS.STB.Analyst
{
    public enum AlarmFlags
    {
        Invalid = 0,
        None = 1,
        Good = 2,
        Alert = 3,
        Danger = 4
    }

    public enum AlarmLevel
    {
        Null = int.MinValue,
        Zero = 0,
        None = 1,
        Good = 2,
        Alert = 3,
        Danger = 4
    }

    public enum AlarmOverallMethod
    {
        None = 0,
        Level = 1,
        InWindow = 2,
        OutOfWindow = 3
    }

    public enum ContainerType
    {
        None,
        Root = 1,
        Set = 2,
        Machine = 3,
        Point = 4
    }

    public enum FreqType
    {
        None = 0,
        Hertz = 1,
        Order = 2
    }

    public enum HierarchyType
    {
        None,
        Hierarchy = 1,
        Route = 2,
        Workspace = 3,
        Template = 4
    }

    public enum MeasurementStatus
    {
        None,
        Active = 0,
        BaseLine = 1,
        ShortTerm = 2,
        LongTerm = 3,
        Unscheduled = 4
    }

    public enum ReadingType
    {
        None,
        Baseline,
        BOV,
        FFT,
        Inspection,
        MCD,
        MeasurementParameter,
        MotorCurrent,
        NonCollection,
        Overall,
        ParametricDigital,
        Phase,
        Time
    }

    public enum SystemDefId
    {
        None,
        Custom = 0,
        Operator = 100,
        Technician = 200,
        Analyst = 300,
        Administrator = 400,
        FieldService = 500
    }

    public enum NoteCategory
    {
        None,
        UserNote,
        CollectionNote,
        NonCollectionNote,
        CodedNote,
        OperatingTimeResetNote,
        AcknowledgeAlarmNote,
        OilAnalysisNote
    }

    public enum MCDParamType
    {
        None = 0,
        Envelope = 1,
        Velocity = 2,
        Temperature = 3
    }

    public enum DetectionType
    {
        None = 0,
        Peak = 20500,
        PkPk = 20501,
        RMS = 20502
    }

    public enum Lines
    {
        lines_none = 0,
        lines_100 = 100,
        lines_200 = 200,
        lines_400 = 400,
        lines_800 = 800,
        lines_1600 = 1600,
        lines_3200 = 3200,
        lines_6400 = 6400,
        lines_12800 = 12800,
        lines_25600 = 25600
    }

    public enum ApplicationId
    {
        None,
        Cyclic,
        General,
        Inspection,
        Lab_Analysis,
        MotorCurrentZoom,
        MotorEnvelopedCurrent,
        OperatingTime,
        Orbit,
        Vibration,
        VibrationDualChannel,
        VibrationTriChannel
    }

    public enum Dad
    {
        None,
        DI1100,
        DerivedPoint,
        DMx,
        Gal,
        Gen,
        IMx,
        IMxM,
        IMxP,
        IMxS,
        IMxT,
        LMU,
        Manual,
        Marlin,
        MasCon,
        Mascon16,
        Microlog,
        MIM,
        OilAnalysis,
        TMU,
        WMx_Sub_WVT,
        WMx
    }

    public enum DadGroup
    {
        None,
        Multilog,
        MicrologAnalyzer,
        MicrologInspector,
        DerivedPoint,
        ManualEntry,
        TrendOil,
        EletronicEntry
    }

    public enum Priority
    {
        None = 0,
        Critical = 1,
        High = 2,
        Medium = 3,
        Low = 4,
        Lowest = 5
    }

    public enum Orientation
    {
        None,
        Horizontal,
        Vertical,
        Axial,
        Radial,
        Triaxial,
        X,
        Y,
        Z
    }

    public enum PointType
    {
        None,
        ACCurrent,
        Acc,
        AccEnvelope,
        AccToDisp,
        AccToVel,
        Battery_Level,
        CountRate,
        Counts,
        Current,
        DerivedOperTimeDays,
        DerivedOperTimeHours,
        DerivedOperTimeMins,
        DerivedOperTimeMonths,
        DerivedOperTimeSecs,
        DerivedOperTimeWeeks,
        DerivedPointCalculated,
        DerivedPointMARLIN,
        Displacement,
        DisplacementMM,
        DualAcc,
        DualAccEnvelope,
        DualAccToDisp,
        DualAccToVel,
        DualDisplacement,
        DualDisplacementMM,
        DualVelEnvelope,
        DualVelToDisp,
        DualVelocity,
        Duration,
        DynamicPressure,
        Efficiency,
        Flow,
        Humidity,
        InternalTemperature,
        LinearDisplacement,
        Logic,
        MCD,
        MicrologCyclicAcceleration,
        MicrologCyclicDisplacement,
        MicrologCyclicEnvelopeAcceleration,
        MicrologCyclicEnvelopeVelocity,
        MicrologCyclicPressure,
        MicrologCyclicSEE,
        MicrologCyclicVelocity,
        MicrologCyclicVolts,
        MicrologMotorCurrentZoom,
        MicrologMotorEnvelopedCurrent,
        MultipleInspection,
        Noise_Level,
        Oil_Analysis,
        OrbitAcc,
        OrbitDisplacement,
        OrbitDisplacementMM,
        OrbitVelocity,
        PeakHFD,
        Power,
        PowerFactor,
        Pressure,
        RMSHFD,
        RPM,
        ResistanceMohms,
        ResistanceOhms,
        SPM,
        Sees,
        Signal_Level,
        Signal_Quality,
        SingleInspection,
        Temperature,
        TransitionalLogic,
        TransitionalSpeed,
        TriAcc,
        TriAccEnvelope,
        TriAccToDisp,
        TriAccToVel,
        TriChannelAcc,
        TriChannelAccEnvelope,
        TriChannelAccToDisp,
        TriChannelAccToVel,
        TriChannelDisplacement,
        TriDisplacement,
        TriVelDisplacement,
        TriVelEnvelope,
        TriVelocity,
        VelEnvelope,
        VelToDisp,
        Velocity,
        VoltsAc,
        VoltsDc,
        Wildcard
    }

    public enum PointTech
    {
        None,
        Accel,
        Velocity,
        Displacement,
        Envelope,
        Temperature,
        Process,
        SEE,
        Wildcard,
        Speed,
        HFD,
        Inspection,
        Time,
        Pressure,
        Volts,
        Current,
        SPM,
        Logic,
        dB,
        Perc,
        OilAnalysis
    }

    public enum Techniques
    {
        All,
        Vibration,
        Sensitive,
        Derivated,
        TrendOil,
        IMx,
        MCD
    }

    public enum Segment
    {
        DS2SAM = 0,
        MachineView = 1,
        ODR = 2
    }
}
