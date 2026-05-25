CREATE TABLE `antropometrias` (
	`id` int AUTO_INCREMENT NOT NULL,
	`userId` int NOT NULL,
	`participantId` varchar(255) NOT NULL,
	`date` timestamp NOT NULL,
	`bracoMeasurements` text NOT NULL,
	`cinturaMeasurements` text NOT NULL,
	`panturrilhaMeasurements` text NOT NULL,
	`createdAt` timestamp NOT NULL DEFAULT (now()),
	`updatedAt` timestamp NOT NULL DEFAULT (now()) ON UPDATE CURRENT_TIMESTAMP,
	CONSTRAINT `antropometrias_id` PRIMARY KEY(`id`)
);
--> statement-breakpoint
CREATE TABLE `fpmEvaluations` (
	`id` int AUTO_INCREMENT NOT NULL,
	`userId` int NOT NULL,
	`participantId` varchar(255) NOT NULL,
	`date` timestamp NOT NULL,
	`dominantHand` varchar(20) NOT NULL,
	`bestLeg` varchar(20) NOT NULL,
	`rightMeasurements` text NOT NULL,
	`leftMeasurements` text NOT NULL,
	`createdAt` timestamp NOT NULL DEFAULT (now()),
	`updatedAt` timestamp NOT NULL DEFAULT (now()) ON UPDATE CURRENT_TIMESTAMP,
	CONSTRAINT `fpmEvaluations_id` PRIMARY KEY(`id`)
);
--> statement-breakpoint
CREATE TABLE `isakEvaluations` (
	`id` int AUTO_INCREMENT NOT NULL,
	`userId` int NOT NULL,
	`participantId` varchar(255) NOT NULL,
	`date` timestamp NOT NULL,
	`measurements` text NOT NULL,
	`createdAt` timestamp NOT NULL DEFAULT (now()),
	`updatedAt` timestamp NOT NULL DEFAULT (now()) ON UPDATE CURRENT_TIMESTAMP,
	CONSTRAINT `isakEvaluations_id` PRIMARY KEY(`id`)
);
--> statement-breakpoint
CREATE TABLE `users` (
	`id` int AUTO_INCREMENT NOT NULL,
	`openId` varchar(64) NOT NULL,
	`name` text,
	`email` varchar(320),
	`loginMethod` varchar(64),
	`role` enum('user','admin') NOT NULL DEFAULT 'user',
	`createdAt` timestamp NOT NULL DEFAULT (now()),
	`updatedAt` timestamp NOT NULL DEFAULT (now()) ON UPDATE CURRENT_TIMESTAMP,
	`lastSignedIn` timestamp NOT NULL DEFAULT (now()),
	CONSTRAINT `users_id` PRIMARY KEY(`id`),
	CONSTRAINT `users_openId_unique` UNIQUE(`openId`)
);
