-- phpMyAdmin SQL Dump
-- version 5.1.3
-- https://www.phpmyadmin.net/
--
-- Host: mysql
-- Erstellungszeit: 13. Apr 2023 um 21:13
-- Server-Version: 10.11.2-MariaDB-1:10.11.2+maria~ubu2204-log
-- PHP-Version: 8.1.17

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Datenbank: `automatenhandel24_connector`
--

-- --------------------------------------------------------

--
-- Tabellenstruktur für Tabelle `access_token`
--

CREATE TABLE `access_token` (
  `id` int(11) NOT NULL,
  `type` varchar(10) NOT NULL,
  `access_token` text NOT NULL,
  `refresh_token` text NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- --------------------------------------------------------

--
-- Tabellenstruktur für Tabelle `error_log`
--

CREATE TABLE `error_log` (
  `id` int(11) NOT NULL,
  `error_type` text NOT NULL,
  `error_module` text NOT NULL,
  `error_message` text NOT NULL,
  `create_date` datetime NOT NULL DEFAULT current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- --------------------------------------------------------

--
-- Tabellenstruktur für Tabelle `lexoffice_hubspot`
--

CREATE TABLE `lexoffice_hubspot` (
  `id` int(11) NOT NULL,
  `deal_id` text NOT NULL,
  `document_id` text NOT NULL,
  `over_lexoffice` int(1) NOT NULL DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- --------------------------------------------------------

--
-- Tabellenstruktur für Tabelle `setting_data`
--

CREATE TABLE `setting_data` (
  `id` int(11) NOT NULL,
  `setting_name` text NOT NULL,
  `setting_value` text NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

-- --------------------------------------------------------

--
-- Tabellenstruktur für Tabelle `users`
--

CREATE TABLE `users` (
  `id` int(11) NOT NULL,
  `username` text NOT NULL,
  `password` text NOT NULL,
  `registred` datetime NOT NULL DEFAULT current_timestamp(),
  `last_login` datetime DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Daten für Tabelle `users`
-- Standard User: admin@domain.com
-- Standard Password: 1234
--

INSERT INTO `users` (`id`, `username`, `password`) VALUES
(1, 'admin@domain.com', '$2b$10$7ur90aQx.sU85TxPmYysAeojzMR.9bVvyW/GBjqxl77BIDyAGHOQ.');

--
-- Indizes der exportierten Tabellen
--

--
-- Indizes für die Tabelle `access_token`
--
ALTER TABLE `access_token`
  ADD PRIMARY KEY (`id`);


--
-- Indizes für die Tabelle `error_log`
--
ALTER TABLE `error_log`
  ADD PRIMARY KEY (`id`);


--
-- Indizes für die Tabelle `lexoffice_hubspot`
--
ALTER TABLE `lexoffice_hubspot`
  ADD PRIMARY KEY (`id`);

--
-- Indizes für die Tabelle `setting_data`
--
ALTER TABLE `setting_data`
  ADD PRIMARY KEY (`id`);

--
-- Indizes für die Tabelle `users`
--
ALTER TABLE `users`
  ADD PRIMARY KEY (`id`);

--
-- AUTO_INCREMENT für exportierte Tabellen
--

--
-- AUTO_INCREMENT für Tabelle `access_token`
--
ALTER TABLE `access_token`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT;


--
-- AUTO_INCREMENT für Tabelle `error_log`
--
ALTER TABLE `error_log`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT;

--
-- AUTO_INCREMENT für Tabelle `setting_data`
--
ALTER TABLE `setting_data`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT;

--
-- AUTO_INCREMENT für Tabelle `users`
--
ALTER TABLE `users`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=2;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
