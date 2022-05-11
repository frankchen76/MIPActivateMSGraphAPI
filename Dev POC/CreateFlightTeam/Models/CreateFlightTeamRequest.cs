﻿// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
using System;

namespace MIPActivateMSGraphAPI_CreateFlightTeam.Models
{
    class CreateFlightTeamRequest
    {
        public int SharePointItemId { get; set; }
        public float FlightNumber { get; set; }
        public string Description { get; set; }
        public string Admin { get; set; }
        public string[] Pilots { get; set; }
        public string[] FlightAttendants { get; set; }
        public string CateringLiaison { get; set; }
        public string DepartureGate { get; set; }
        public DateTimeOffset DepartureTime { get; set; }
    }
}
